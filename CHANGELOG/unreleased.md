# Unreleased

These are the changes that will go out in the next release.

## Changed

- **Breaking:** `SharePointClient` now takes a `CertificateCredential` instead of a `get_token`
  callable. Authentication runs entirely in Rust, so no Python code executes on the request path.
- Token refresh moved to a background task owned by the client. It refreshes five minutes before
  expiry and on any 401 or 403, and entering the async context manager waits for the first token so
  invalid credentials fail at connect time.
- The token scope is now derived from the site host rather than supplied by the caller.
- Replaced the Python runtime implementation with a PyO3 extension using Tokio and Reqwest.
- Preserved the public async client method names and signatures.
- Switched package builds from hatchling to maturin and removed the unused CLI metadata.
- Published type information through `async_sharepoint.pyi`.

## Added

- `CertificateCredential` for the Entra ID certificate client-credentials flow.
- `SharePointClient.from_static_token()` for wrapping a token you already hold.

Async Sharepoint started out as a project generated from [Cookiecutter PyPackage](https://github.com/audreyfeldroy/cookiecutter-pypackage) containing:

- Initial scaffold for Async Sharepoint.
- `src/async_sharepoint/` package with CLI (Typer + Rich), py.typed marker
- Tests with pytest, coverage across Python 3.12/3.13/3.14
- CI via GitHub Actions: lint (Ruff), type check (ty), test matrix, coverage reporting
- Security scanning: CodeQL analysis for public repositories, Dependabot, and a Zizmor workflow audit
- Docs site with Zensical + mkdocstrings and a GitHub Pages deployment workflow
- Trusted publishing to PyPI with OIDC and build provenance attestation
- `justfile` with dev commands: qa, test, type-check, docs-serve, release
- Issue templates, PR template, contributing guide, code of conduct, security policy
- MIT license, .editorconfig, .gitignore

### Contributors

[@deanm0000](https://github.com/deanm0000) (Dean MacGregor) created Async Sharepoint.
