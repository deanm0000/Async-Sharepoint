# Usage

To use Async Sharepoint in a project:

```python
from async_sharepoint import SharePointClient


def get_token() -> dict:
    return token_provider.acquire_token_for_client(scopes=[scope])


async with SharePointClient("https://example.sharepoint.com/sites/Team", get_token) as client:
    documents = await client.get_items("Documents")
    content = await documents[0].download()
```

The token callback is synchronous and must return a mapping containing `access_token` and `expires_in`.
When `max_wait` is supplied to a paginated operation, the result is a pair containing the items already
received and an awaitable for the remaining items.
