# Async Sharepoint

[![PyPI version](https://img.shields.io/pypi/v/Async-Sharepoint.svg)](https://pypi.org/project/Async-Sharepoint/)
[![PyPI downloads](https://static.pepy.tech/badge/Async-Sharepoint/month)](https://pepy.tech/projects/Async-Sharepoint)

An async SharePoint client with a native Rust implementation exposed through PyO3.

The repository also contains a standalone Rust crate in `core/`, published as
`async-sharepoint`, for Rust projects that do not need Python bindings.

* [GitHub](https://github.com/deanm0000/Async-Sharepoint/) | [PyPI](https://pypi.org/project/Async-Sharepoint/) | [Documentation](https://deanm0000.github.io/Async-Sharepoint/)
* Created by [Dean MacGregor](na) | GitHub [@deanm0000](https://github.com/deanm0000) | PyPI [@deanm0000](https://pypi.org/user/deanm0000/)
* MIT License

## Features

* Async list, item, file, folder, permission, and search operations
* Native Entra ID certificate authentication with background token refresh
* Automatic token caching and refresh
* Retry handling for throttling and transient SharePoint failures
* Paginated queries and chunked uploads
* Python 3.10+ stable-ABI Linux wheels

## Installation

```bash
uv add Async-Sharepoint
## or
pip install Async-Sharepoint
```

### Rust

```bash
cargo add async-sharepoint
```

```rust
use async_sharepoint::{CertificateCredential, SharePointClient};

let credential = CertificateCredential::load(
    tenant_id,
    client_id,
    "/path/to/private.key",
    thumbprint,
)?;
let client = SharePointClient::new(site_url, &credential)?;
let files = client.ls(None).await?;
```

## Usage

```python
from async_sharepoint import CertificateCredential, SharePointClient

credential = CertificateCredential(
    tenant_id=tenant_id,
    client_id=client_id,
    private_key_path="/path/to/private.key",
    thumbprint=thumbprint,
)

async with SharePointClient(site_url, credential) as client:
    documents = await client.get_items("Documents")
```

## Documentation

Full documentation is available on
[GitHub Pages](https://deanm0000.github.io/Async-Sharepoint/).

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, testing, and
documentation instructions.

