import asyncio
import random
import time
import uuid
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import overload
from urllib.parse import parse_qsl, quote, urlencode, urlparse, urlsplit, urlunsplit

import httpx

UPLOAD_CHUNK_SIZE = 4 * 1024 * 1024
# refresh this far ahead of actual expiry
TOKEN_REFRESH_MARGIN = 60.0
# floor on cached token lifetime, so a missing/short "expires_in" can't spin the refresh loop
MIN_TOKEN_LIFETIME = 60.0
TOKEN_REFRESH_TIMEOUT = 30.0


def _lit(value: str) -> str:
    """Builds a quoted+escaped OData string literal for use inside method-call parens."""
    escaped = (
        str(value).replace("%", "%25").replace("+", "%2B").replace("#", "%23").replace("&", "%26").replace("'", "''")
    )
    return "'" + escaped + "'"


def _bool_lit(value: bool) -> str:
    return "true" if value else "false"


def _extract_search_rows(payload: dict) -> list[dict]:
    result = payload.get("PrimaryQueryResult")
    if result is None:
        result = payload.get("d", {}).get("query", {}).get("PrimaryQueryResult")

    rows = result.get("RelevantResults", {}).get("Table", {}).get("Rows", []) if result else []

    normalized_rows = []
    for row in rows:
        cells = row.get("Cells", [])
        if isinstance(cells, dict):
            cells = cells.get("results", [])
        normalized_rows.append({cell.get("Key"): cell.get("Value") for cell in cells})
    return normalized_rows


class _PropertiesAttrMixin:
    """Falls back to `self.properties[name]` for attribute access not otherwise defined."""

    properties: dict

    def __getattr__(self, name: str):
        try:
            return self.properties[name]
        except KeyError:
            raise AttributeError(name) from None


@dataclass
class SPFolder(_PropertiesAttrMixin):
    client: "SharePointClient | None" = None
    server_relative_url: str | None = None
    properties: dict = field(default_factory=dict)
    list_url: str | None = None
    item_id: str | None = None


@dataclass
class SPFile(_PropertiesAttrMixin):
    client: "SharePointClient"
    server_relative_path: str | None = None
    properties: dict = field(default_factory=dict)
    list_url: str | None = None
    item_id: str | None = None
    resolved: bool = True
    _resolve_task: "asyncio.Task[None] | None" = field(default=None, repr=False, compare=False)

    @classmethod
    def _from_json(cls, client: "SharePointClient", data: dict) -> "SPFile":
        path = data.get("ServerRelativeUrl", "")
        return cls(client=client, server_relative_path=path, properties=data)

    def _start_resolve(self) -> None:
        """Kicks off a background fetch of this list item's File metadata."""
        self.resolved = False
        self._resolve_task = asyncio.create_task(self._resolve())

    async def _resolve(self) -> None:
        url = f"{self.list_url}/items({self.item_id})/File"
        data = await self.client._get_json(url)
        self.properties = {**self.properties, **data}
        self.server_relative_path = data.get("ServerRelativeUrl", self.server_relative_path)
        self.resolved = True

    async def resolve(self) -> None:
        """Resolve this file's metadata immediately, pre-empting any background resolve.

        Cancels the background resolve task started when this file was returned
        from a list-items lookup, if it hasn't finished yet, then fetches now.
        """
        if self.resolved:
            return
        if self._resolve_task is not None and not self._resolve_task.done():
            self._resolve_task.cancel()
        await self._resolve()

    async def download(self) -> bytes:
        """Download this file's raw content, resolving its metadata first if needed.

        Returns
        -------
        bytes
            The file's raw content.

        Raises
        ------
        RuntimeError
            If the file has no server-relative path even after resolving.
        """
        if not self.resolved:
            await self.resolve()
        if self.server_relative_path is None:
            raise RuntimeError("file has no server_relative_path even after resolving")
        return await self.client.download(self.server_relative_path)

    async def get_list_item(self) -> "SPItem":
        """Fetch the list item backing this file, resolving this file first if needed.

        Returns
        -------
        SPFile or SPFolder or SharePointClient
            The list item behind this file (typically another `SPFile`).

        Raises
        ------
        RuntimeError
            If the file has no server-relative path even after resolving.
        """
        if not self.resolved:
            await self.resolve()
        if self.server_relative_path is None:
            raise RuntimeError("file has no server_relative_path even after resolving")
        url = (
            f"{self.client._web_url}/GetFileByServerRelativePath"
            f"(DecodedUrl={_lit(self.server_relative_path)})/ListItemAllFields"
        )
        data = await self.client._get_json(url)
        list_id = data.get("ParentList", {}).get("Id", "")
        list_url = f"{self.client._web_url}/lists/GetById({_lit(list_id)})"
        return self.client._item_from_json(list_url, data)

    def get_url(self) -> str:
        if not self.resolved:
            raise RuntimeError("file is not resolved run await resolve() first")
        if self.server_relative_path is None:
            raise RuntimeError("file has no server_relative_path even after resolving")
        parsed = urlsplit(self.properties["ServerRedirectedEmbedUri"])
        query = parse_qsl(parsed.query, keep_blank_values=True)
        new_query = [(k, v if k != "action" else "default") for k, v in query]

        return urlunsplit(
            (
                parsed.scheme,
                parsed.netloc,
                quote(parsed.path, safe="/"),
                urlencode(new_query),
                parsed.fragment,
            )
        )


@dataclass
class SPList(_PropertiesAttrMixin):
    client: "SharePointClient"
    id: str
    title: str
    properties: dict

    @classmethod
    def _from_json(cls, client: "SharePointClient", data: dict) -> "SPList":
        return cls(
            client=client,
            id=data.get("Id", ""),
            title=data.get("Title", ""),
            properties=data,
        )

    async def get_items(self, *, caml: str | None = None, max_wait: float | None = None):
        """Fetch items belonging to this list.

        Parameters
        ----------
        caml : str, optional
            CAML query ViewXml to filter/scope the items.
        max_wait : float, optional
            If given, pages for at most this many seconds before returning,
            handing off remaining pages as a background task.

        Returns
        -------
        list[SPItem] or tuple[list[SPItem], asyncio.Task[list[SPItem]]]
            See `SharePointClient.get_items`.
        """
        return await self.client.get_items(id=self.id, caml=caml, max_wait=max_wait)


class SharePointClient(_PropertiesAttrMixin):
    """Async client for a SharePoint site's REST API.

    Parameters
    ----------
    site_url : str
        Root URL of the SharePoint site, e.g. "https://contoso.sharepoint.com/sites/Team".
    get_token : Callable[[], dict]
        Synchronous callable returning a token payload with at least "access_token"
        and "expires_in".
    properties : dict, optional
        Raw list-item properties, when this client represents a site-link item
        returned from `get_items`.
    item_id : str, optional
        The originating list item's id, when this client represents a site-link item.
    list_url : str, optional
        The originating list's URL, when this client represents a site-link item.
    """

    def __init__(
        self,
        site_url: str,
        get_token: Callable[[], dict],
        *,
        properties: dict | None = None,
        item_id: str | None = None,
        list_url: str | None = None,
    ):
        self.site_url = site_url.rstrip("/")
        self._web_url = f"{self.site_url}/_api/web"
        self._get_token = get_token
        self._client = httpx.AsyncClient(timeout=30.0)
        self._token: str | None = None
        self._token_expiry = 0.0
        self._token_lock = asyncio.Lock()
        self._token_task: asyncio.Task[None] | None = None
        self._refresh_task: asyncio.Task[None] | None = None
        self._loop: asyncio.AbstractEventLoop | None = None
        # only populated when this client represents a site-link item returned from get_items
        self.properties = properties or {}
        self.item_id = item_id
        self.list_url = list_url

    async def aclose(self) -> None:
        """Cancel the background token-refresh task and close the underlying HTTP client."""
        if self._refresh_task is not None:
            self._refresh_task.cancel()
        if self._token_task is not None:
            self._token_task.cancel()
        await self._client.aclose()

    async def __aenter__(self) -> "SharePointClient":
        return self

    async def __aexit__(self, *exc_info) -> None:
        await self.aclose()

    # -- token / http plumbing -------------------------------------------------

    async def _acquire_token(self) -> None:
        token_data = await asyncio.to_thread(self._get_token)
        self._token = token_data["access_token"]
        try:
            lifetime = float(token_data.get("expires_in") or 0.0)
        except (TypeError, ValueError):
            lifetime = 0.0
        self._token_expiry = time.monotonic() + max(lifetime - TOKEN_REFRESH_MARGIN, MIN_TOKEN_LIFETIME)

    async def _refresh_token(self) -> None:
        if self._token is not None and time.monotonic() < self._token_expiry:
            return
        try:
            await asyncio.wait_for(self._token_lock.acquire(), timeout=TOKEN_REFRESH_TIMEOUT)
        except TimeoutError:
            raise TimeoutError("timed out waiting for token refresh lock") from None
        try:
            if self._token is not None and time.monotonic() < self._token_expiry:
                return
            if self._token_task is None or self._token_task.done():
                self._token_task = asyncio.create_task(self._acquire_token())
            token_task = self._token_task
        finally:
            self._token_lock.release()

        try:
            await asyncio.wait_for(asyncio.shield(token_task), timeout=TOKEN_REFRESH_TIMEOUT)
        except TimeoutError:
            raise TimeoutError("timed out acquiring a SharePoint access token") from None

    async def _token_refresh_loop(self) -> None:
        """Proactively refreshes the token ahead of expiry; backs off and retries on transient failures."""
        while True:
            try:
                await asyncio.sleep(max(self._token_expiry - time.monotonic(), MIN_TOKEN_LIFETIME))
                await self._refresh_token()
            except asyncio.CancelledError:
                return
            except Exception:
                await asyncio.sleep(5)

    def _reset_loop_state(self) -> None:
        """Drops lock/task state bound to a previous event loop (e.g. a notebook's per-cell asyncio.run)."""
        loop = asyncio.get_running_loop()
        if self._loop is loop:
            return
        if self._refresh_task is not None:
            self._refresh_task.cancel()
            self._refresh_task = None
        if self._token_task is not None:
            self._token_task.cancel()
            self._token_task = None
        self._token_lock = asyncio.Lock()
        self._loop = loop

    async def _headers(self) -> dict:
        self._reset_loop_state()
        if self._token is None or time.monotonic() >= self._token_expiry:
            await self._refresh_token()
        if self._refresh_task is None or self._refresh_task.done():
            self._refresh_task = asyncio.create_task(self._token_refresh_loop())
        return {
            "Authorization": f"Bearer {self._token}",
            "Accept": "application/json;odata=nometadata",
        }

    async def _request(
        self,
        method: str,
        url: str,
        *,
        params: dict | None = None,
        json_body: dict | None = None,
        content: bytes | None = None,
        max_retries: int = 6,
    ) -> httpx.Response:
        did_force_refresh = False
        for attempt in range(max_retries + 1):
            headers = await self._headers()
            response = await self._client.request(
                method,
                url,
                params=params,
                json=json_body,
                content=content,
                headers=headers,
            )
            if response.status_code in (401, 403):
                if did_force_refresh:
                    response.raise_for_status()
                did_force_refresh = True
                self._token_expiry = 0.0
                continue
            if response.status_code not in (429, 503, 504):
                response.raise_for_status()
                return response
            if attempt == max_retries:
                response.raise_for_status()

            retry_after = response.headers.get("Retry-After")
            try:
                delay = float(retry_after) if retry_after else None
            except ValueError:
                delay = None
            if delay is None:
                delay = min(1.0 * (2**attempt), 30.0)
            await asyncio.sleep(delay + random.uniform(0, 0.5))

        raise RuntimeError("retry loop exited unexpectedly")

    async def _get_json(self, url: str, *, params: dict | None = None) -> dict:
        response = await self._request("GET", url, params=params)
        return response.json()

    async def _post_json(
        self,
        url: str,
        *,
        json_body: dict | None = None,
        content: bytes | None = None,
    ) -> dict:
        response = await self._request("POST", url, json_body=json_body, content=content)
        if not response.content:
            return {}
        return response.json()

    async def _get_all(self, url: str, *, params: dict | None = None) -> list[dict]:
        results: list[dict] = []
        next_url, next_params = url, params
        while next_url:
            data = await self._get_json(next_url, params=next_params)
            results.extend(data.get("value", []))
            next_url = data.get("odata.nextLink")
            next_params = None
        return results

    async def _continue_paging(self, next_url: str | None) -> list[dict]:
        """Fetches whatever pages remain, starting at next_url (or none if already exhausted)."""
        results: list[dict] = []
        while next_url:
            data = await self._get_json(next_url)
            results.extend(data.get("value", []))
            next_url = data.get("odata.nextLink")
        return results

    async def _get_all_deferred(
        self, url: str, *, params: dict | None = None, max_wait: float
    ) -> tuple[list[dict], "asyncio.Task[list[dict]]"]:
        """Pages until max_wait elapses, then hands off the rest as a background task."""
        start = time.monotonic()
        data = await self._get_json(url, params=params)
        results = list(data.get("value", []))
        next_url = data.get("odata.nextLink")

        while next_url and time.monotonic() - start <= max_wait:
            data = await self._get_json(next_url)
            results.extend(data.get("value", []))
            next_url = data.get("odata.nextLink")

        return results, asyncio.create_task(self._continue_paging(next_url))

    # -- lists -------------------------------------------------------------

    @overload
    async def get(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        max_wait: None = None,
    ) -> "SPList | list[SPList]": ...
    @overload
    async def get(
        self, title: None = None, *, id: None = None, max_wait: float
    ) -> tuple[list[SPList], "asyncio.Task[list[SPList]]"]: ...
    async def get(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        max_wait: float | None = None,
    ):
        """Fetch a single list by title or id, or all lists in the site.

        Parameters
        ----------
        title : str, optional
            Title of the list to fetch. Mutually exclusive with `id`.
        id : str, optional
            Id of the list to fetch. Mutually exclusive with `title`.
        max_wait : float, optional
            Only valid when fetching all lists. If given, pages for at most this
            many seconds before returning, handing off remaining pages as a
            background task.

        Returns
        -------
        SPList or list[SPList] or tuple[list[SPList], asyncio.Task[list[SPList]]]
            A single `SPList` when `title`/`id` is given; otherwise a list of all
            lists, or `(partial_list, background_task)` when `max_wait` is given.

        Raises
        ------
        ValueError
            If both `title` and `id` are given, or if `max_wait` is given together
            with `title`/`id`.
        """
        if title is not None and id is not None:
            raise ValueError("specify only one of title or id")
        if title is not None or id is not None:
            if max_wait is not None:
                raise ValueError("max_wait is only supported when fetching all lists")
            if title is not None:
                data = await self._get_json(f"{self._web_url}/lists/GetByTitle({_lit(title)})")
            elif id is not None:
                data = await self._get_json(f"{self._web_url}/lists/GetById({_lit(id)})")
            else:
                raise ValueError("s/b impossible")
            return SPList._from_json(self, data)

        if max_wait is None:
            raw_lists = await self._get_all(f"{self._web_url}/lists")
            return [SPList._from_json(self, d) for d in raw_lists]

        raw_lists, raw_task = await self._get_all_deferred(f"{self._web_url}/lists", max_wait=max_wait)
        lists_so_far = [SPList._from_json(self, d) for d in raw_lists]

        async def _map_remaining() -> list[SPList]:
            return [SPList._from_json(self, d) for d in await raw_task]

        return lists_so_far, asyncio.create_task(_map_remaining())

    async def get_default_document_library(self) -> SPList:
        """Fetch the site's default document library.

        Returns
        -------
        SPList
            The default document library.
        """
        data = await self._get_json(f"{self._web_url}/DefaultDocumentLibrary")
        return SPList._from_json(self, data)

    # -- items ---------------------------------------------------------------

    def _item_from_json(self, list_url: str, data: dict) -> "SPFile | SPFolder | SharePointClient":
        item_id = str(data.get("Id"))
        link_url = (data.get("Link") or {}).get("Url")
        if link_url:
            return SharePointClient(link_url, self._get_token, properties=data, item_id=item_id, list_url=list_url)
        if data.get("FileSystemObjectType") == 1:
            return SPFolder(
                client=self,
                server_relative_url=data.get("ServerRelativeUrl") or data.get("FileRef"),
                properties=data,
                list_url=list_url,
                item_id=item_id,
            )

        file = SPFile(client=self, list_url=list_url, item_id=item_id, properties=data)
        file._start_resolve()
        return file

    @overload
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        max_wait: None = None,
    ) -> "list[SPItem]": ...
    @overload
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        max_wait: float,
    ) -> "tuple[list[SPItem], asyncio.Task[list[SPItem]]]": ...
    async def get_items(
        self,
        title: str | None = None,
        *,
        id: str | None = None,
        caml: str | None = None,
        max_wait: float | None = None,
    ):
        """Fetch items belonging to a list, by title or id.

        Parameters
        ----------
        title : str, optional
            Title of the parent list. Mutually exclusive with `id`.
        id : str, optional
            Id of the parent list. Mutually exclusive with `title`.
        caml : str, optional
            CAML query ViewXml to filter/scope the items. When omitted, all items
            are fetched.
        max_wait : float, optional
            If given, pages for at most this many seconds before returning,
            handing off remaining pages as a background task.

        Returns
        -------
        list[SPItem] or tuple[list[SPItem], asyncio.Task[list[SPItem]]]
            A list of `SPFile`/`SPFolder`/`SharePointClient` items, or
            `(partial_list, background_task)` when `max_wait` is given.

        Raises
        ------
        ValueError
            If both `title` and `id` are given, or if neither is given.
        """
        if title is not None and id is not None:
            raise ValueError("specify only one of title or id")
        if title is not None:
            list_url = f"{self._web_url}/lists/GetByTitle({_lit(title)})"
        elif id is not None:
            list_url = f"{self._web_url}/lists/GetById({_lit(id)})"
        else:
            raise ValueError("must specify title or id")

        if caml is not None:
            body = {
                "query": {"__metadata": {"type": "SP.CamlQuery"}, "ViewXml": caml},
            }
            data = await self._post_json(f"{list_url}/GetItems", json_body=body)
            items = [self._item_from_json(list_url, d) for d in data.get("value", [])]
            if max_wait is None:
                return items
            # GetItems isn't paginated here, so the background task has nothing left to do.
            return items, asyncio.create_task(self._continue_paging(None))

        if max_wait is None:
            raw_items = await self._get_all(f"{list_url}/items")
            return [self._item_from_json(list_url, d) for d in raw_items]

        raw_items, raw_task = await self._get_all_deferred(f"{list_url}/items", max_wait=max_wait)
        items = [self._item_from_json(list_url, d) for d in raw_items]

        async def _map_remaining() -> list[SPItem]:
            return [self._item_from_json(list_url, d) for d in await raw_task]

        return items, asyncio.create_task(_map_remaining())

    # -- files -----------------------------------------------------------------

    async def get_file(self, path: str) -> SPFile:
        """Reference a file by its server-relative path without making a request.

        Parameters
        ----------
        path : str
            Server-relative path of the file, e.g. "/sites/Team/Shared Documents/file.txt".

        Returns
        -------
        SPFile
            A file reference that can be downloaded or otherwise resolved on demand.
        """
        return SPFile(client=self, server_relative_path=path)

    async def download(self, path: str) -> bytes:
        """Download a file's raw content by server-relative path.

        Parameters
        ----------
        path : str
            Server-relative path of the file.

        Returns
        -------
        bytes
            The file's raw content.
        """
        url = f"{self._web_url}/GetFileByServerRelativePath(DecodedUrl={_lit(path)})/$value"
        response = await self._request("GET", url)
        return response.content

    async def upload(self, folder_path: str, filename: str, content: bytes, *, overwrite: bool = True) -> SPFile:
        """Upload content to a folder, chunking automatically for large payloads.

        Parameters
        ----------
        folder_path : str
            Server-relative path of the destination folder.
        filename : str
            Name to give the uploaded file.
        content : bytes
            File content to upload.
        overwrite : bool, optional
            Whether to overwrite an existing file of the same name (default True).

        Returns
        -------
        SPFile
            The uploaded file.
        """
        folder_url = f"{self._web_url}/GetFolderByServerRelativeUrl({_lit(folder_path)})"
        add_url = f"{folder_url}/Files/add(url={_lit(filename)},overwrite={_bool_lit(overwrite)})"

        if len(content) <= UPLOAD_CHUNK_SIZE:
            data = await self._post_json(add_url, content=content)
            return SPFile._from_json(self, data)

        await self._post_json(add_url, content=b"")
        file_url = f"{folder_url}/Files({_lit(filename)})"
        upload_id = str(uuid.uuid4())

        pos = 0
        first_chunk = content[pos : pos + UPLOAD_CHUNK_SIZE]
        pos += len(first_chunk)
        await self._post_json(f"{file_url}/startUpload(uploadID={_lit(upload_id)})", content=first_chunk)

        while True:
            chunk = content[pos : pos + UPLOAD_CHUNK_SIZE]
            if pos + len(chunk) < len(content):
                await self._post_json(
                    f"{file_url}/continueUpload(uploadID={_lit(upload_id)},fileOffset={pos})",
                    content=chunk,
                )
                pos += len(chunk)
            else:
                data = await self._post_json(
                    f"{file_url}/finishUpload(uploadID={_lit(upload_id)},fileOffset={pos})",
                    content=chunk,
                )
                break

        return SPFile._from_json(self, data)

    # -- folders / users / permissions -----------------------------------------

    async def add_folder(self, path: str, *, overwrite: bool = False) -> SPFolder:
        """Create a folder at a server-relative path.

        Parameters
        ----------
        path : str
            Server-relative path of the folder to create.
        overwrite : bool, optional
            Whether to overwrite an existing folder at that path (default False).

        Returns
        -------
        SPFolder
            The created folder.
        """
        url = f"{self._web_url}/Folders/AddUsingPath(DecodedUrl={_lit(path)},Overwrite={_bool_lit(overwrite)})"
        data = await self._post_json(url)
        return SPFolder(client=self, server_relative_url=data.get("ServerRelativeUrl", path), properties=data)

    async def get_current_user(self) -> dict:
        """Fetch the currently authenticated user.

        Returns
        -------
        dict
            The raw `CurrentUser` payload.
        """
        return await self._get_json(f"{self._web_url}/CurrentUser")

    async def get_effective_permissions(self, path: str, *, login_name: str | None = None) -> dict:
        """Fetch a user's effective permissions on a folder path.

        Parameters
        ----------
        path : str
            Server-relative path of the folder to check permissions on.
        login_name : str, optional
            Login name of the user to check. Defaults to the current authenticated user.

        Returns
        -------
        dict
            The raw effective-permissions payload.
        """
        if login_name is None:
            current_user = await self.get_current_user()
            login_name = str(current_user["LoginName"])
        url = (
            f"{self._web_url}/GetFolderByServerRelativeUrl({_lit(path)})/ListItemAllFields"
            f"/GetUserEffectivePermissions({_lit(login_name)})"
        )
        return await self._post_json(url)

    # -- search ------------------------------------------------------------

    async def search(
        self,
        text: str,
        *,
        title: str | None = None,
        id: str | None = None,
        row_limit: int = 6,
    ) -> list[dict]:
        """Run a scoped search query against a specific list.

        Parameters
        ----------
        text : str
            Search text to match.
        title : str, optional
            Title of the list to scope the search to. Mutually exclusive with `id`.
        id : str, optional
            Id of the list to scope the search to. Mutually exclusive with `title`.
        row_limit : int, optional
            Maximum number of results to return (default 6).

        Returns
        -------
        list[dict]
            Normalized result rows keyed by column name.

        Raises
        ------
        ValueError
            If neither `title` nor `id` is given.
        """
        if title is None and id is None:
            raise ValueError("must specify title or id")
        sp_list = await self.get(title, id=id)
        assert isinstance(sp_list, SPList)
        site = await self._get_json(f"{self.site_url}/_api/site")
        web = await self._get_json(self._web_url)
        root_folder = await self._get_json(f"{self._web_url}/lists/GetById({_lit(sp_list.id)})/RootFolder")

        site_id = site["Id"]
        web_id = web["Id"]
        root_folder_url = root_folder["ServerRelativeUrl"]

        parsed_site_url = urlparse(self.site_url)
        origin = f"{parsed_site_url.scheme}://{parsed_site_url.netloc}"
        root_folder_abs_url = origin + quote(root_folder_url, safe="/%")
        parent_link_url = origin + root_folder_url

        query_template = (
            "{searchTerms} "
            f"(siteId:{{{site_id}}} OR siteId:{site_id}) "
            f"(webId:{{{web_id}}} OR webId:{web_id}) "
            f"(NormListID:{sp_list.id}) "
            f'(path:"{root_folder_abs_url}" OR ParentLink:"{parent_link_url}*") '
            "ContentTypeId:0x0*"
        )
        select_properties = [
            "editorowsuser",
            "authorowsuser",
            "Filename",
            "SPSiteURL",
            "Title",
            "ParentLink",
            "ListItemID",
            "ListID",
            "contentclass",
            "IsDocument",
            "IsContainer",
            "FileExtension",
            "SecondaryFileExtension",
            "OriginalPath",
            "DefaultEncodingURL",
            "ServerRedirectedURL",
            "ServerRedirectedPreviewURL",
            "LastModifiedTime",
            "SharedWithUsersOWSUser",
            "HitHighlightedSummary",
            "ModifierDates",
            "LastModifiedTimeForRetention",
        ]
        params = {
            "querytext": f"'({text}*)'",
            "querytemplate": f"'{query_template}'",
            "selectproperties": "'" + ",".join(select_properties) + "'",
            "SummaryLength": "100",
            "RowLimit": str(row_limit),
            "culture": "1033",
            "BypassResultTypes": "true",
            "EnableQueryRules": "false",
            "ProcessBestBets": "false",
            "ProcessPersonalFavorites": "false",
            "clienttype": "'sug_SPListInline'",
            "properties": "'EnableDynamicGroups:true'",
            "TrimDuplicates": "false",
        }

        response = await self._request("GET", f"{self.site_url}/_api/search/query", params=params)
        return _extract_search_rows(response.json())


# a list item is a file, a folder, or a link to another SharePoint site
SPItem = SPFile | SPFolder | SharePointClient

__all__ = ["SPItem", "SPFile", "SPFolder", "SharePointClient"]
