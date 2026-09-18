# API Reference

The implementation is a native PyO3 extension. Type annotations are shipped in
`async_sharepoint.pyi`.

## CertificateCredential

```python
CertificateCredential(
  *,
  tenant_id,
  client_id,
  private_key_path,
  thumbprint,
)
```

Entra ID certificate credentials. The private key and thumbprint are parsed on construction, so
malformed input raises immediately. Read-only attributes: `tenant_id`, `client_id`.

## SharePointClient

```python
SharePointClient(
  site_url,
  credential,
  *,
  properties=None,
  item_id=None,
  list_url=None,
)

SharePointClient.from_static_token(
  site_url,
  token,
  *,
  properties=None,
  item_id=None,
  list_url=None,
)
```

`from_static_token` wraps a token you already hold; it is never refreshed.

The client is an async context manager. Its public methods are:

- `await aclose()`
- `await get(title=None, *, id=None, max_wait=None)`
- `await get_default_document_library()`
- `await get_items()`
- `await get_file(path)`
- `await ls()`
- `await download(path)`

`path` accepts a server-relative path, a bare or `{braced}` `UniqueId`, an `AllItems.aspx?id=...`
browser link, or a `Doc.aspx?sourcedoc={GUID}` browser/share link (the latter two are also
resolved via the file's `UniqueId`).

- `await download_chunks(path)` returns an async context manager exposing `await get_chunk()`,
  which streams the file via HTTP `Range` requests and returns `None` once fully read.
- `await download_file(path, local_path)` downloads directly to `local_path` in chunks
  without buffering the whole file in memory. Returns `None`.
- `await upload(full_path, content, *, overwrite=True)`
- `upload_chunks(full_path, *, overwrite=True)` returns an async context manager exposing
  `await write(chunk)` for each chunk of the upload, in order.
- `await upload_file(full_path, local_path, *, overwrite=True)` uploads a local file in chunks
  without buffering it in memory. Returns `None`.
- `await add_folder(path, *, overwrite=False)`
- `await get_current_user()`
- `await get_effective_permissions(path, *, login_name=None)`
- `await search(text, *, title=None, id=None, row_limit=6)`

## SPFile

File values expose `client`, `properties`,
`list_url`, `item_id`, `unique_id`, and `resolved` (a read-only property that is true once
resolution has completed). Their public methods are `resolve`, `download`, `download_chunks`,
`download_file`, `get_url`, `browser_url`, and `embed_url`. `download_chunks()` and
`download_file(local_path)` mirror the client's methods but need no `path` argument since the
file is already resolved.

## SPFolder

Folder values expose `client`, `server_relative_url`, `properties`, `list_url`,
`item_id`, and `resolved`. Their public methods are `resolve`, `ls`, and `get_url`.

## SPList

List values are returned by `SharePointClient.get` and
`get_default_document_library`. They expose `client`, `id`, `title`,
`properties`, `get_items(*, caml=None, folder_path=None, max_wait=None)`, and `get_url`.
