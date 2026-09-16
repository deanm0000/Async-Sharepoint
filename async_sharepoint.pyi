from collections.abc import Awaitable
from typing import Any, TypeAlias, overload

class CertificateCredential:
    @property
    def tenant_id(self) -> str: ...
    @property
    def client_id(self) -> str: ...
    def __init__(
        self,
        *,
        tenant_id: str,
        client_id: str,
        private_key_path: str,
        thumbprint: str,
    ) -> None: ...

class SPFolder:
    client: SharePointClient | None
    server_relative_url: str | None
    properties: dict[str, Any]
    list_url: str | None
    item_id: str | None
    @property
    def resolved(self) -> bool: ...
    def __init__(
        self,
        client: SharePointClient | None = None,
        server_relative_url: str | None = None,
        properties: dict[str, Any] | None = None,
        list_url: str | None = None,
        item_id: str | None = None,
    ) -> None: ...
    async def resolve(self) -> None: ...
    async def ls(self) -> list[SPItem]: ...
    def get_url(self) -> str: ...
    def __getattr__(self, name: str) -> Any: ...

class SPFile:
    client: SharePointClient
    server_relative_path: str | None
    properties: dict[str, Any]
    list_url: str | None
    item_id: str | None
    @property
    def resolved(self) -> bool: ...
    def __init__(
        self,
        client: SharePointClient,
        server_relative_path: str | None = None,
        properties: dict[str, Any] | None = None,
        list_url: str | None = None,
        item_id: str | None = None,
    ) -> None: ...
    async def resolve(self) -> None: ...
    async def download(self) -> bytes: ...
    def get_url(self) -> str: ...
    def browser_url(self) -> str: ...
    def __getattr__(self, name: str) -> Any: ...

class SPList:
    client: SharePointClient
    id: str
    title: str
    properties: dict[str, Any]

    @overload
    async def get_items(
        self,
        *,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: None = None,
    ) -> list[SPItem]: ...
    @overload
    async def get_items(
        self,
        *,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: float,
    ) -> tuple[list[SPItem], Awaitable[list[SPItem]]]: ...
    def get_url(self) -> str: ...
    def __getattr__(self, name: str) -> Any: ...

class SharePointClient:
    site_url: str
    properties: dict[str, Any]
    item_id: str | None
    list_url: str | None

    def __init__(
        self,
        site_url: str,
        credential: CertificateCredential,
        *,
        properties: dict[str, Any] | None = None,
        item_id: str | None = None,
        list_url: str | None = None,
    ) -> None: ...
    @staticmethod
    def from_static_token(
        site_url: str,
        token: str,
        *,
        properties: dict[str, Any] | None = None,
        item_id: str | None = None,
        list_url: str | None = None,
    ) -> SharePointClient: ...
    async def aclose(self) -> None: ...
    async def __aenter__(self) -> SharePointClient: ...
    async def __aexit__(self, *exc_info: object) -> None: ...
    @overload
    async def get(self, title: str, *, id: None = None, max_wait: None = None) -> SPList: ...
    @overload
    async def get(self, title: None = None, *, id: str, max_wait: None = None) -> SPList: ...
    @overload
    async def get(self, title: None = None, *, id: None = None, max_wait: None = None) -> list[SPList]: ...
    @overload
    async def get(
        self, title: None = None, *, id: None = None, max_wait: float
    ) -> tuple[list[SPList], Awaitable[list[SPList]]]: ...
    async def get_default_document_library(self) -> SPList: ...
    @overload
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: None = None,
    ) -> list[SPItem]: ...
    @overload
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: float,
    ) -> tuple[list[SPItem], Awaitable[list[SPItem]]]: ...
    async def get_file(self, path: str) -> SPFile: ...
    async def ls(self, path: str | None = None) -> list[SPItem]: ...
    async def download(self, path: str) -> bytes: ...
    async def upload(
        self,
        folder_path: str,
        filename: str,
        content: bytes,
        *,
        overwrite: bool = True,
    ) -> SPFile: ...
    async def add_folder(self, path: str, *, overwrite: bool = False) -> SPFolder: ...
    async def get_current_user(self) -> dict[str, Any]: ...
    async def get_effective_permissions(self, path: str, *, login_name: str | None = None) -> dict[str, Any]: ...
    async def search(
        self,
        text: str,
        *,
        title: str | None = None,
        id: str | None = None,
        row_limit: int = 6,
    ) -> list[dict[str, Any]]: ...
    def __getattr__(self, name: str) -> Any: ...

SPItem: TypeAlias = SPFile | SPFolder | SharePointClient
__all__: list[str]
