from asyncio import Task
from collections.abc import Awaitable
from typing import Any, TypeAlias, overload
from uuid import UUID

class CertificateCredential:
    """Entra ID certificate credentials for SharePoint.

    Parameters
    ----------
    tenant_id : str
        Entra tenant identifier.
    client_id : str
        Entra application client identifier.
    private_key_path : str
        Path to a PEM private key or combined certificate-and-key file.
    thumbprint : str
        Hex-encoded SHA-1 certificate thumbprint.

    Examples
    --------
    >>> credential = CertificateCredential(
    ...     tenant_id="00000000-0000-0000-0000-000000000000",
    ...     client_id="11111111-1111-1111-1111-111111111111",
    ...     private_key_path="/path/to/private.key",
    ...     thumbprint="D33CFD3BB83E0EFB90AF709C897025286244FA0E",
    ... )
    """

    @property
    def tenant_id(self) -> str:
        """Entra tenant identifier."""
        ...
    @property
    def client_id(self) -> str:
        """Entra application client identifier."""
        ...
    def __init__(
        self,
        *,
        tenant_id: str,
        client_id: str,
        private_key_path: str,
        thumbprint: str,
    ) -> None: ...

class SPFolder:
    """A SharePoint folder returned by a list or folder operation.

    Attributes
    ----------
    client : SharePointClient or None
        Client used to resolve and access the folder.
    server_relative_url : str or None
        Server-relative folder URL, available after resolution.
    properties : dict[str, Any]
        SharePoint properties. Keys are also available through attribute access.
    list_url : str or None
        REST URL of the folder's parent list.
    item_id : str or None
        SharePoint list-item identifier.
    resolved : bool
        Whether background property resolution has finished.
    """

    client: SharePointClient | None
    server_relative_url: str | None
    properties: dict[str, Any]
    list_url: str | None
    item_id: str | None
    @property
    def resolved(self) -> bool:
        """Whether background property resolution has finished."""
        ...
    def __init__(
        self,
        client: SharePointClient | None = None,
        server_relative_url: str | None = None,
        properties: dict[str, Any] | None = None,
        list_url: str | None = None,
        item_id: str | None = None,
    ) -> None: ...
    async def resolve(self) -> None:
        """Wait for the folder's background property resolution.

        This is usually not necessary unless you need to ensure all properties are available before accessing them.
        Examples
        --------
        >>> await folder.resolve()
        >>> folder.Name
        'Reports'
        """
        ...
    async def ls(self) -> list[SPItem]:
        """List the files and folders directly inside this folder.

        Returns
        -------
        list[SPItem]
            Child files, folders, and linked SharePoint sites.

        Examples
        --------
        >>> items = await folder.ls()
        """
        ...
    async def del_folder(self, *, ignore_missing: bool = False, recursive: bool = False) -> None:
        """Delete this SharePoint folder.

        The folder must be empty unless ``recursive=True``. Set
        ``ignore_missing=True`` to return successfully if it no longer exists.
        """
        ...
    def get_url(self) -> str:
        """Return a SharePoint browser URL for the folder.

        Returns
        -------
        str
            URL that opens the folder in its document library.

        Examples
        --------
        >>> folder_url = folder.get_url()
        """
        ...
    def __getattr__(self, name: str) -> Any: ...
    def __repr__(self) -> str: ...
    def __str__(self) -> str: ...

class SPFile:
    """A SharePoint file returned by a file or list operation.

    Attributes
    ----------
    client : SharePointClient
        Client used to resolve and access the file.
    properties : dict[str, Any]
        SharePoint properties. Keys are
        also available through attribute access.
    list_url : str or None
        REST URL of the file's parent list.
    item_id : str or None
        SharePoint list-item identifier.
    unique_id : str or None
        SharePoint file GUID when the file was located by a sharing URL.
    resolved : bool
        Whether background property resolution has finished.
    """

    client: SharePointClient
    properties: dict[str, Any]
    list_url: str | None
    item_id: str | None
    unique_id: str | None
    @property
    def resolved(self) -> bool:
        """Whether background property resolution has finished."""
        ...
    def __init__(
        self,
        client: SharePointClient,
        server_relative_path: str | None = None,
        properties: dict[str, Any] | None = None,
        list_url: str | None = None,
        item_id: str | None = None,
        unique_id: str | None = None,
    ) -> None: ...
    async def resolve(self) -> None:
        """Wait for the file's background property resolution.

        Examples
        --------
        >>> await file.resolve()
        >>> file.Name
        'report.xlsx'
        """
        ...
    async def download(self) -> bytes:
        """Download the file contents.

        Returns
        -------
        bytes
            Complete binary file content.

        Examples
        --------
        >>> content = await file.download()
        """
        ...
    def download_chunks(self) -> DownloadChunks:
        """Open an async context manager that streams this file's contents in chunks.

        Returns
        -------
        DownloadChunks
            Async context manager exposing ``get_chunk()``.

        Examples
        --------
        >>> async with file.download_chunks() as download:
        ...     while (chunk := await download.get_chunk()) is not None:
        ...         handle(chunk)
        """
        ...
    async def download_file(self, local_path: str) -> None:
        """Download the file to a local path, streaming it instead of buffering it in memory.

        Parameters
        ----------
        local_path : str
            Local filesystem path to write the file to.

        Examples
        --------
        >>> await file.download_file("/tmp/report.xlsx")
        """
        ...
    async def del_file(self) -> None:
        """Delete this SharePoint file."""
        ...
    def get_url(self) -> str:
        """Return SharePoint's preferred URL for the file.

        Returns
        -------
        str
            URL that opens the file with the default action.

        Examples
        --------
        >>> await file.resolve()
        >>> file_url = file.get_url()
        """
        ...
    def browser_url(self) -> str:
        """Return the file's document-library browser URL.

        Returns
        -------
        str
            ``AllItems.aspx`` URL selecting the file.

        Examples
        --------
        >>> await file.resolve()
        >>> file_url = file.browser_url()
        """
        ...
    def embed_url(self) -> str:
        """Return a URL suitable for embedding or previewing the file.

        Uses ``ServerRedirectedEmbedUri`` if present, otherwise builds a Doc.aspx
        preview link from the file's GUID (``UniqueId``, else ``ContentTag``/``ETag``).

        Returns
        -------
        str
            URL that opens an interactive preview of the file.

        Examples
        --------
        >>> await file.resolve()
        >>> preview_url = file.embed_url()
        """
        ...
    def __getattr__(self, name: str) -> Any: ...
    def __repr__(self) -> str: ...
    def __str__(self) -> str: ...

class SPList:
    """A SharePoint list.

    Attributes
    ----------
    client : SharePointClient
        Client used to access the list.
    id : str
        SharePoint list identifier.
    title : str
        Display title of the list.
    properties : dict[str, Any]
        SharePoint list properties. Keys are also available through attribute access.
    """

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
    ) -> list[SPItem]:
        """Get items from the list.

        Parameters
        ----------
        caml : str, optional
            CAML view XML used to query the list.
        folder_path : str, optional
            Server-relative folder path used to scope a CAML query.
        max_wait : float, optional
            Initial pagination budget in seconds. Supplying it changes the
            return value to an initial list and an awaitable for the remainder.

        Returns
        -------
        list[SPItem] or tuple[list[SPItem], Awaitable[list[SPItem]]]
            All items, or the initial and remaining items when ``max_wait`` is supplied.

        Examples
        --------
        >>> items = await documents.get_items()
        >>> items, remaining = await documents.get_items(max_wait=1)
        >>> all_items = items + await remaining
        """
        ...
    @overload
    async def get_items(
        self,
        *,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: float,
    ) -> tuple[list[SPItem], Task[list[SPItem]]]: ...
    def get_url(self) -> str:
        """Return the list's SharePoint browser URL.

        Returns
        -------
        str
            URL of the list's ``AllItems.aspx`` view.

        Examples
        --------
        >>> list_url = documents.get_url()
        """
        ...
    def __getattr__(self, name: str) -> Any: ...

class UploadChunks:
    """Async context manager returned by ``SharePointClient.upload_chunks``.

    Examples
    --------
    >>> async with client.upload_chunks("/sites/Team/Documents/report.txt") as upload:
    ...     await upload.write(chunk_one)
    ...     await upload.write(chunk_two)
    """

    async def __aenter__(self) -> UploadChunks: ...
    async def __aexit__(self, *exc_info: object) -> None: ...
    async def write(self, chunk: bytes) -> None:
        """Write the next chunk of the file being uploaded.

        Parameters
        ----------
        chunk : bytes
            Next chunk of binary file content, in order.
        """
        ...

class DownloadChunks:
    """Async context manager returned by ``SharePointClient.download_chunks`` and
    ``SPFile.download_chunks``.

    Examples
    --------
    >>> async with client.download_chunks("/sites/Team/Documents/report.txt") as download:
    ...     while (chunk := await download.get_chunk()) is not None:
    ...         handle(chunk)
    """

    async def __aenter__(self) -> DownloadChunks: ...
    async def __aexit__(self, *exc_info: object) -> None: ...
    async def get_chunk(self) -> bytes | None:
        """Fetch the next chunk of the file's contents.

        Returns
        -------
        bytes or None
            Next chunk of binary file content, or ``None`` once the file has been
            fully read.
        """
        ...

class SharePointClient:
    """Asynchronous client for a SharePoint site.

    Parameters
    ----------
    site_url : str
        Absolute URL of the SharePoint site.
    credential : CertificateCredential
        Certificate credentials used to acquire and refresh access tokens.
    properties : dict[str, Any], optional
        Properties attached when this object represents a linked SharePoint site.
    item_id : str, optional
        Parent list-item identifier for a linked site.
    list_url : str, optional
        Parent list REST URL for a linked site.

    Attributes
    ----------
    site_url : str
        Normalized SharePoint site URL.
    properties : dict[str, Any]
        Properties attached to the site object.
    item_id : str or None
        Parent list-item identifier for a linked site.
    list_url : str or None
        Parent list REST URL for a linked site.

    Examples
    --------
    >>> async with SharePointClient(site_url, credential) as client:
    ...     items = await client.ls()
    """

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
    ) -> SharePointClient:
        """Create a client using an existing access token.

        Parameters
        ----------
        site_url : str
            Absolute URL of the SharePoint site.
        token : str
            Existing bearer token. The client does not refresh it.
        properties : dict[str, Any], optional
            Properties attached to the client.
        item_id : str, optional
            Parent list-item identifier.
        list_url : str, optional
            Parent list REST URL.

        Returns
        -------
        SharePointClient
            Client backed by the supplied static token.

        Examples
        --------
        >>> client = SharePointClient.from_static_token(site_url, access_token)
        """
        ...
    async def aclose(self) -> None:
        """Close the client when it is not managed by an async context.

        Examples
        --------
        >>> await client.aclose()
        """
        ...
    async def __aenter__(self) -> SharePointClient:
        """Wait for authentication and enter the async context."""
        ...
    async def __aexit__(self, *exc_info: object) -> None: ...
    @overload
    async def get(self, title: str, *, id: None = None, max_wait: None = None) -> SPList:
        """Get SharePoint lists.

        Parameters
        ----------
        title : str, optional
            Title of one list to return.
        id : str, optional
            Identifier of one list to return. Do not combine with ``title``.
        max_wait : float, optional
            Initial pagination budget in seconds when fetching all lists.

        Returns
        -------
        SPList or list[SPList] or tuple[list[SPList], Awaitable[list[SPList]]]
            One selected list, all lists, or initial and remaining lists when
            ``max_wait`` is supplied.

        Examples
        --------
        >>> lists = await client.get()
        >>> documents = await client.get("Documents")
        >>> lists, remaining = await client.get(max_wait=1)
        """
        ...
    @overload
    async def get(self, title: None = None, *, id: str, max_wait: None = None) -> SPList: ...
    @overload
    async def get(self, title: None = None, *, id: None = None, max_wait: None = None) -> list[SPList]: ...
    @overload
    async def get(
        self, title: None = None, *, id: None = None, max_wait: float
    ) -> tuple[list[SPList], Awaitable[list[SPList]]]: ...
    async def get_default_document_library(self) -> SPList:
        """Get the site's default document library.

        Returns
        -------
        SPList
            Default document library.

        Examples
        --------
        >>> documents = await client.get_default_document_library()
        """
        ...
    @overload
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        folder_path: str | None = None,
        max_wait: None = None,
    ) -> list[SPItem]:
        """Get items from a list selected by title or identifier.

        Parameters
        ----------
        title : str, optional
            Title of the list to query.
        id : str, optional
            Identifier of the list to query. Do not combine with ``title``.
        caml : str, optional
            CAML view XML used to query the list.
        folder_path : str, optional
            Server-relative folder path used to scope a CAML query.
        max_wait : float, optional
            Initial pagination budget in seconds. Supplying it changes the
            return value to an initial list and an awaitable for the remainder.

        Returns
        -------
        list[SPItem] or tuple[list[SPItem], Awaitable[list[SPItem]]]
            All items, or initial and remaining items when ``max_wait`` is supplied.

        Examples
        --------
        >>> items = await client.get_items("Documents")
        >>> items, remaining = await client.get_items("Documents", max_wait=1)
        """
        ...
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
    async def get_file(self, path: str | UUID) -> SPFile:
        """Get a resolved SharePoint file.

        Parameters
        ----------
        path : str | UUID
            Server-relative path, bare/braced UniqueId, ``AllItems.aspx?id=...`` URL, or
            ``Doc.aspx?sourcedoc=...`` URL.

        Returns
        -------
        SPFile
            Resolved file object.

        Examples
        --------
        >>> file = await client.get_file("/sites/Team/Documents/report.xlsx")
        """
        ...
    async def ls(self, path: str | None = None) -> list[SPItem]:
        """List a folder in the default document library.

        Parameters
        ----------
        path : str, optional
            Server-relative folder path. The document-library root is used when omitted.

        Returns
        -------
        list[SPItem]
            Files and folders directly inside the selected folder.

        Examples
        --------
        >>> root_items = await client.ls()
        """
        ...
    async def download(self, path: str | UUID) -> bytes:
        """Download a SharePoint file by path or browser URL.

        Parameters
        ----------
        path : str | UUID
            Server-relative path, bare/braced UniqueId, ``AllItems.aspx?id=...`` URL, or
            ``Doc.aspx?sourcedoc=...`` URL.

        Returns
        -------
        bytes
            Complete binary file content.

        Examples
        --------
        >>> content = await client.download("/sites/Team/Documents/report.xlsx")
        """
        ...
    async def upload(
        self,
        full_path: str,
        content: bytes,
        *,
        overwrite: bool = True,
    ) -> SPFile:
        """Upload a file to SharePoint, creating missing parent folders.

        Parameters
        ----------
        full_path : str
            Server-relative destination path, including the filename.
        content : bytes
            Binary file content.
        overwrite : bool, optional
            Whether to replace an existing file. Defaults to ``True``.

        Returns
        -------
        SPFile
            Uploaded file.

        Examples
        --------
        >>> uploaded = await client.upload(
        ...     "/sites/Team/Documents/report.txt",
        ...     b"report contents",
        ... )
        """
        ...
    def upload_chunks(self, full_path: str, *, overwrite: bool = True) -> UploadChunks:
        """Upload chunks to SharePoint, creating missing parent folders.

        Parameters
        ----------
        full_path : str
            Server-relative destination path, including the filename.
        overwrite : bool, optional
            Whether to replace an existing file. Defaults to ``True``.

        Returns
        -------
        UploadChunks
            Async context manager exposing ``write(chunk)``.

        Examples
        --------
        >>> async with client.upload_chunks("/sites/Team/Documents/report.txt") as upload:
        ...     await upload.write(b"first chunk")
        ...     await upload.write(b"last chunk")
        """
        ...
    async def upload_file(self, full_path: str, local_path: str, *, overwrite: bool = True) -> None:
        """Stream a local file to SharePoint, creating missing parent folders.

        Parameters
        ----------
        full_path : str
            Server-relative destination path, including the filename.
        local_path : str
            Local filesystem path of the file to upload.
        overwrite : bool, optional
            Whether to replace an existing file. Defaults to ``True``.

        Examples
        --------
        >>> await client.upload_file("/sites/Team/Documents/report.xlsx", "/tmp/report.xlsx")
        """
        ...
    def download_chunks(self, path: str | UUID) -> DownloadChunks:
        """Open an async context manager that streams a file's contents in chunks.

        Parameters
        ----------
        path : str | UUID
            Server-relative path, bare/braced UniqueId, ``AllItems.aspx?id=...`` URL, or
            ``Doc.aspx?sourcedoc=...`` URL.

        Returns
        -------
        DownloadChunks
            Async context manager exposing ``get_chunk()``.

        Examples
        --------
        >>> async with client.download_chunks("/sites/Team/Documents/report.xlsx") as download:
        ...     while (chunk := await download.get_chunk()) is not None:
        ...         handle(chunk)
        """
        ...
    async def download_file(self, path: str | UUID, local_path: str) -> None:
        """Download a SharePoint file to a local path, streaming it instead of buffering it in memory.

        Parameters
        ----------
        path : str | UUID
            Server-relative path, bare/braced UniqueId, ``AllItems.aspx?id=...`` URL, or
            ``Doc.aspx?sourcedoc=...`` URL.
        local_path : str
            Local filesystem path to write the file to.

        Examples
        --------
        >>> await client.download_file("/sites/Team/Documents/report.xlsx", "/tmp/report.xlsx")
        """
        ...
    async def add_folder(self, path: str, *, overwrite: bool = False) -> SPFolder:
        """Create a SharePoint folder.

        Parameters
        ----------
        path : str
            Server-relative path of the folder to create.
        overwrite : bool, optional
            Whether an existing folder is accepted. Defaults to ``False``.

        Returns
        -------
        SPFolder
            Created folder.

        Examples
        --------
        >>> folder = await client.add_folder("/sites/Team/Documents/Reports")
        """
        ...
    async def del_file(self, path: str | UUID) -> None:
        """Delete a SharePoint file by path, browser URL, or UniqueId."""
        ...
    async def del_folder(
        self,
        path: str,
        *,
        ignore_missing: bool = False,
        recursive: bool = False,
    ) -> None:
        """Delete a SharePoint folder.

        Parameters
        ----------
        path : str
            Server-relative folder path or ``AllItems.aspx?id=...`` URL.
        ignore_missing : bool, optional
            Return successfully when the folder does not exist.
        recursive : bool, optional
            Delete the folder and all descendants. By default, only empty
            folders can be deleted.
        """
        ...
    async def get_current_user(self) -> dict[str, Any]:
        """Get the current SharePoint user.

        Returns
        -------
        dict[str, Any]
            SharePoint user properties.

        Examples
        --------
        >>> user = await client.get_current_user()
        """
        ...
    async def get_effective_permissions(self, path: str, *, login_name: str | None = None) -> dict[str, Any]:
        """Get effective permissions for a SharePoint path.

        Parameters
        ----------
        path : str
            Server-relative path whose permissions are requested.
        login_name : str, optional
            SharePoint login name. The current user is used when omitted.

        Returns
        -------
        dict[str, Any]
            SharePoint permission mask fields.

        Examples
        --------
        >>> permissions = await client.get_effective_permissions("/sites/Team/Documents")
        """
        ...
    async def search(
        self,
        text: str,
        *,
        title: str | None = None,
        id: str | None = None,
        row_limit: int = 6,
    ) -> list[dict[str, Any]]:
        """Search a SharePoint list.

        Parameters
        ----------
        text : str
            Text to search for.
        title : str, optional
            Title of the list to search.
        id : str, optional
            Identifier of the list to search. Do not combine with ``title``.
        row_limit : int, optional
            Maximum number of search rows. Defaults to 6.

        Returns
        -------
        list[dict[str, Any]]
            Search result rows.

        Examples
        --------
        >>> results = await client.search("quarterly forecast", title="Documents")
        """
        ...
    def __getattr__(self, name: str) -> Any: ...

SPItem: TypeAlias = SPFile | SPFolder | SharePointClient
__all__: list[str]
