# Async Sharepoint

[![PyPI version](https://img.shields.io/pypi/v/Async-Sharepoint.svg)](https://pypi.org/project/Async-Sharepoint/)
[![PyPI downloads](https://static.pepy.tech/badge/Async-Sharepoint/month)](https://pepy.tech/projects/Async-Sharepoint)

An async SharePoint client with a native Rust implementation exposed through PyO3.

* [GitHub](https://github.com/deanm0000/Async-Sharepoint/) | [PyPI](https://pypi.org/project/Async-Sharepoint/) | [Documentation](https://deanm0000.github.io/Async-Sharepoint/)
* Created by [Dean MacGregor](na) | GitHub [@deanm0000](https://github.com/deanm0000) | PyPI [@deanm0000](https://pypi.org/user/deanm0000/)
* MIT License

## Features

* Async list, item, file, folder, permission, and search operations
* Automatic token caching and refresh
* Retry handling for throttling and transient SharePoint failures
* Paginated queries and chunked uploads
* Python 3.10+ stable-ABI Linux wheels

## Installation

```bash
uv add Async-Sharepoint
```

## Usage

```python
from async_sharepoint import SharePointClient

async with SharePointClient(site_url, get_token) as client:
    documents = await client.get_items("Documents")
```

## Documentation

Full documentation is available on
[GitHub Pages](https://deanm0000.github.io/Async-Sharepoint/).

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, testing, and
documentation instructions.

## Author

Async Sharepoint was created in 2026 by Dean MacGregor.

Built with [Cookiecutter](https://github.com/cookiecutter/cookiecutter) and the [audreyfeldroy/cookiecutter-pypackage](https://github.com/audreyfeldroy/cookiecutter-pypackage) project template.
