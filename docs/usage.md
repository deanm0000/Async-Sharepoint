# Usage

To use Async Sharepoint in a project:

```python
from async_sharepoint import CertificateCredential, SharePointClient

credential = CertificateCredential(
    tenant_id="00000000-0000-0000-0000-000000000000",
    client_id="11111111-1111-1111-1111-111111111111",
    private_key_path="/path/to/azure-app-private.key",
    thumbprint="447A9DA9F1B534BC0956CDAA3AD6D6C1E436B8D3",
)

async with SharePointClient("https://example.sharepoint.com/sites/Team", credential) as client:
    documents = await client.get_items("Documents")
    content = await documents[0].download()
```

`CertificateCredential` performs the Entra ID client-credentials flow in Rust. The private key file
may be a standalone PEM key or a combined certificate-and-key file; only the `PRIVATE KEY` block is
read. The thumbprint is the certificate's hex-encoded SHA-1 digest, as shown in the Azure portal.

The scope is derived from the site host, so `https://example.sharepoint.com/sites/Team` requests
`https://example.sharepoint.com/.default`.

A background task owned by the client refreshes the token five minutes before it expires and on any
401 or 403 response. Entering the async context manager waits for the first token, so invalid
credentials raise immediately rather than on the first request.

To use a token you already hold, build the client with `SharePointClient.from_static_token(site_url,
token)`. That token is never refreshed.

When `max_wait` is supplied to a paginated operation, the result is a pair containing the items already
received and an awaitable for the remaining items.
