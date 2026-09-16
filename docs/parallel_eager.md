# Eager and parallel fetching

SharePoint list queries often involve two independent sources of latency: walking every page of a
large result and fetching the file or folder properties omitted from each list-item response.
Async Sharepoint overlaps both kinds of work.

## Optimistically parallel pagination

With the default `max_wait=None`, `get()` and `get_items()` return a complete result. After reading
the first page, the client examines SharePoint's continuation token and optimistically predicts the
tokens for later pages. It requests the real continuation chain and predicted pages concurrently,
deduplicates overlapping results, restores item order, and cancels guesses beyond the final page.

This is optimistic because SharePoint item IDs can have gaps, and a page can advance farther than
predicted. The client continuously adjusts its frontier from real continuation links. If a
continuation token cannot be interpreted safely, it falls back to ordinary sequential pagination.

## Eager item refinement

A SharePoint list-item response identifies whether an item is a file or folder, but it does not
contain every useful property. Other clients commonly make you fetch each item's `File` or `Folder`
object manually, one at a time, before fields such as the file name and server-relative path are
available.

Async Sharepoint creates the appropriate `SPFile` or `SPFolder` immediately and starts its property
request in the background. Every item's refinement runs concurrently, so a list with many items
does not pay that extra network latency serially.

```python
items = await client.get_items("Documents")
files = [item for item in items if item.FileSystemObjectType == 0]

# Refinement began as soon as each SPFile was created.
await files[0].resolve()
print(files[0].Name, files[0].server_relative_path)
```

You usually do not need to call `resolve()` yourself. Async methods that need refined data, such as
`SPFile.download()` and `SPFolder.ls()`, await the existing background task. Call `resolve()` when
you need to guarantee that synchronous property access or URL generation is ready. You can also check if it's resolved with `item.resolved`

## Start work early with `max_wait`

Sometimes receiving the first useful batch quickly matters more than receiving the complete result
in one await. Pass `max_wait` to put a time budget, in seconds, on the initial sequential page walk:

```python
items, remaining = await client.get_items("Documents", max_wait=1)

for item in items:
    process_first_batch(item)

all_items = items + await remaining
```

The first page is always fetched. The client follows additional pages until the budget is exceeded,
then returns:

1. The items fetched so far, whose property refinement has already started.
2. An awaitable that fetches all remaining pages using optimistic parallel pagination and starts
   refinement for those items as they are materialized.

The budget is checked between page requests; it is not a network timeout and does not cancel a
request already in flight.

Await the second value exactly once and concatenate it with the first list when complete ordering is
needed. The same pattern works for fetching all site lists:

```python
lists, remaining = await client.get(max_wait=1)
all_lists = lists + await remaining
```

`max_wait` applies only to paginated list retrieval. Selecting one list by title or ID already makes
a single request, so `client.get("Documents", max_wait=1)` is rejected.

## Compared with a synchronous query chain

The difference is broader than replacing blocking calls with `await`. Async Sharepoint overlaps
page retrieval where continuation tokens permit it and starts the extra property requests for all
returned items without requiring a manual query per file. See the complete side-by-side download
example in [Usage](usage.md#compared-with-office365-rest-python-client).