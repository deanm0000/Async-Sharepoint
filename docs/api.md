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
- `await get_items(title=None, *, id=None, caml=None, max_wait=None)`
- `await get_file(path)`
- `await download(path)`
- `await upload(folder_path, filename, content, *, overwrite=True)`
- `await add_folder(path, *, overwrite=False)`
- `await get_current_user()`
- `await get_effective_permissions(path, *, login_name=None)`
- `await search(text, *, title=None, id=None, row_limit=6)`

## SPFile

File values expose `client`, `server_relative_path`, `properties`, `list_url`,
`item_id`, and `resolved` (a read-only property that is true once resolution
has completed). Their public methods are `resolve`, `download`, and `get_url`.

## SPFolder

Folder values expose `client`, `server_relative_url`, `properties`, `list_url`,
and `item_id`.

## SPList

List values are returned by `SharePointClient.get` and
`get_default_document_library`. They expose `client`, `id`, `title`,
`properties`, and `get_items(*, caml=None, max_wait=None)`.
