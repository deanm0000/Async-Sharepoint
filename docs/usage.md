# Usage

## Project scope

Async Sharepoint is purpose-built for fast, non-blocking list, file, and folder workflows. It does
not aim to expose every feature in SharePoint. For applications that need broad coverage of the
full SharePoint API, use an established general-purpose SharePoint library instead.

Its narrower focus allows file listings to use [optimistically parallel pagination and eager item
refinement](parallel_eager.md), avoiding a blocking request chain and a separate sequential property
request for every returned file.

That said, I'm not opposed to making it feature complete, it just isn't a priority.

## Connect

Create one client for the site and use it as an async context manager:

```python
from async_sharepoint import CertificateCredential, SharePointClient

credential = CertificateCredential(
    tenant_id="00000000-0000-0000-0000-000000000000",
    client_id="11111111-1111-1111-1111-111111111111",
    private_key_path="/path/to/azure-app-private.key",
    thumbprint="D33CFD3BB83E0EFB90AF709C897025286244FA0E",
)

async with SharePointClient("https://example.sharepoint.com/sites/Team", credential) as client:
    items = await client.ls()
```

The context manager verifies the credentials before entering and releases the client when finished.
For credential setup, static tokens, and token refresh behavior, see [Authentication](auth.md).
If a context manager does not fit your application lifecycle, close the client explicitly:

```python
client = SharePointClient("https://example.sharepoint.com/sites/Team", credential)
try:
    items = await client.ls()
finally:
    await client.aclose()
```

## Lists and list items

Get all lists, one list by title or ID, or the site's default document library:

```python
lists = await client.get()
documents = await client.get("Documents")
documents_by_id = await client.get(id="b8d3d4a2-...")
default_documents = await client.get_default_document_library()

print(documents.title, documents.id)
print(documents.get_url())
```

Get list items through an `SPList`, or directly through the client by title or ID:

```python
items = await documents.get_items()
items_by_title = await client.get_items("Documents")
items_by_id = await client.get_items(id=documents.id)
```

Both forms accept CAML. Use `folder_path` to scope a CAML query to a folder:

```python
caml = """
<View>
  <Query><Where><Eq>
    <FieldRef Name="Status" />
    <Value Type="Text">Active</Value>
  </Eq></Where></Query>
</View>
"""

active_items = await documents.get_items(
    caml=caml,
    folder_path="/sites/Team/Documents/Current",
)
```

### Return promptly with `max_wait`

For a paginated list or item query, `max_wait` limits how long the initial call spends following
pages. The call returns the items collected during that interval and an awaitable for everything
remaining:

```python
items, remaining = await documents.get_items(max_wait=1)

# Start useful work once the initial pagination budget is spent.
for item in items:
    print(item)

all_items = items + await remaining
```

The first page is always fetched. `max_wait` is also available when fetching all lists:

```python
lists, remaining_lists = await client.get(max_wait=1)
all_lists = lists + await remaining_lists
```

See [Eager and parallel fetching](parallel_eager.md) for how pagination and item properties are
fetched concurrently.

## Browse files and folders

`ls()` lists the default document library root, a server-relative folder path, or a folder item:

```python
root_items = await client.ls()
report_items = await client.ls("/sites/Team/Documents/Reports")

report_folder = next(item for item in root_items if item.FileSystemObjectType == 1)
folder_items = await report_folder.ls()
print(report_folder.get_url())
```

Items expose their SharePoint fields through both `properties` and attribute access, so
`item.properties["Name"]` and `item.Name` refer to the same value.

## Get and download files

Get a file using a server-relative path:

```python
file = await client.get_file("/sites/Team/Documents/Reports/summary.xlsx")
content = await file.download()

with open("summary.xlsx", "wb") as output:
    output.write(content)
```

Or download directly when you do not need the `SPFile`:

```python
content = await client.download("/sites/Team/Documents/Reports/summary.xlsx")
```

Both `get_file()` and `download()` also accept native SharePoint browser links. Supported links
include an `AllItems.aspx?id=...` URL and a `Doc.aspx?sourcedoc=...` sharing URL:

```python
browser_link = (
    "https://example.sharepoint.com/sites/Team/Documents/Forms/AllItems.aspx"
    "?id=%2Fsites%2FTeam%2FDocuments%2FReports%2Fsummary.xlsx"
)
file = await client.get_file(browser_link)

sharing_link = (
    "https://example.sharepoint.com/sites/Team/_layouts/15/Doc.aspx"
    "?sourcedoc=%7B01246A4B-84D7-49D6-8937-895D3C0F50A9%7D"
)
content = await client.download(sharing_link)
```

A file's `UniqueId` (the same UUID found in a `sourcedoc` link) is also accepted directly, bare
or `{braced}`:

```python
file = await client.get_file("01246A4B-84D7-49D6-8937-895D3C0F50A9")
file = await client.get_file("{01246A4B-84D7-49D6-8937-895D3C0F50A9}")
file = await client.get_file(UUID("01246A4B-84D7-49D6-8937-895D3C0F50A9"))
```

An `SPFile` can produce three useful SharePoint links. `get_url()` uses SharePoint's preferred URL
for the file, `browser_url()` builds its document-library browsing URL, and `embed_url()` builds a
preview/embed link:

```python
await file.resolve()
preferred_link = file.get_url()
library_link = file.browser_url()
preview_link = file.embed_url()
```

Items returned by a list query start resolving their additional properties immediately in the
background. Methods such as `download()` wait for that work when necessary; call `resolve()`
explicitly before synchronously reading a property or generating a link that depends on it.

### Streaming downloads

For large files, avoid buffering the whole file in memory by downloading directly to disk:

```python
await client.download_file("/sites/Team/Documents/Reports/summary.xlsx", "summary.xlsx")

# Or, from an already-resolved SPFile, no path argument is needed:
await file.download_file("summary.xlsx")
```

Both take chunks over HTTP `Range` requests internally. To process the file chunk-by-chunk
yourself instead, use the async context manager and its `get_chunk()` method, which returns
`None` once the file has been fully read:

```python
async with client.download_chunks("/sites/Team/Documents/Reports/summary.xlsx") as download:
    while (chunk := await download.get_chunk()) is not None:
        handle(chunk)

# Or, from an already-resolved SPFile:
async with file.download_chunks() as download:
    while (chunk := await download.get_chunk()) is not None:
        handle(chunk)
```

## Upload files and create folders

```python
folder = await client.add_folder("/sites/Team/Documents/Exports")
print(folder.get_url())

with open("results.csv", "rb") as source:
    uploaded = await client.upload(
        "/sites/Team/Documents/Exports/results.csv",
        source.read(),
        overwrite=True,
    )

print(uploaded.properties["ServerRelativeUrl"])
```

Set `overwrite=False` on `upload()` to reject an existing filename. `add_folder()` also accepts
`overwrite=True` when an existing folder should be accepted. `upload()`, `upload_file()`, and
`upload_chunks()` automatically create any missing parent folders in the destination path.

Delete files from either the client or an `SPFile`:

```python
await client.del_file("/sites/Team/Documents/Exports/results.csv")
await uploaded.del_file()
```

Folder deletion is non-recursive by default and rejects a nonempty folder. Pass `recursive=True`
to delete all descendants, or `ignore_missing=True` when an absent folder is acceptable:

```python
await folder.del_folder()
await client.del_folder("/sites/Team/Documents/Exports", recursive=True)
await client.del_folder("/sites/Team/Documents/AlreadyGone", ignore_missing=True)
```

### Streaming uploads

To upload a local file without buffering it in memory, use `upload_file()`:

```python
await client.upload_file("/sites/Team/Documents/Exports/results.csv", "results.csv")
```

To write chunks yourself instead, use the async context manager returned by `upload_chunks()`
and call `write()` once per chunk, in order:

```python
async with client.upload_chunks("/sites/Team/Documents/Exports/results.csv") as upload:
    for chunk in chunks:
        await upload.write(chunk)
```

## Search, users, and permissions

Search within a list selected by title or ID:

```python
results = await client.search("quarterly forecast", title="Documents", row_limit=20)
results_by_id = await client.search("quarterly forecast", id=documents.id)
```

Get the current SharePoint user and effective permissions for either that user or a specified
login:

```python
current_user = await client.get_current_user()
my_permissions = await client.get_effective_permissions("/sites/Team/Documents")
user_permissions = await client.get_effective_permissions(
    "/sites/Team/Documents",
    login_name="user@example.com",
)
```

## Compared with Office365-REST-Python-Client

A basic file download with `office365-rest-python-client` is synchronous and requires each query
to be executed explicitly. A list item also needs another file query before it knows the file's
name:

```python
ctx = ClientContext(some_site).with_access_token(gen_token)
ctx_all = ctx.web.lists.get_all().execute_query()
docs = ctx_all[8]
items = docs.get_items().execute_query()
first_file = next(x for x in items if x.properties.get("FileSystemObjectType") == 0)
usable_file = first_file.file.get().execute_query()
file_name = usable_file.properties.get("Name")
with open(target_path, "wb") as output:
    usable_file.download(output).execute_query()
```

Async Sharepoint performs I/O asynchronously and begins resolving every returned item's additional
properties in parallel:

```python
ctx = SharePointClient(some_site, credential)
root = await ctx.get_default_document_library()
items, more_items = await root.get_items(max_wait=1)
usable_file = next(x for x in items if x.FileSystemObjectType == 0)
await usable_file.resolve()
file_name = usable_file.Name
with open(target_path, "wb") as output:
    output.write(await usable_file.download())
```

Unlike the synchronous query chain, `resolve()` joins property retrieval that was already started
in parallel for every item. Await `more_items` when the operation needs the rest of the paginated
result.
