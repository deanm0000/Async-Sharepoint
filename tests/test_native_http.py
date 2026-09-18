import json
import threading
from collections.abc import Iterator
from contextlib import contextmanager
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import unquote, urlsplit

import pytest

from async_sharepoint import SharePointClient, SPFile, SPFolder


class SharePointHandler(BaseHTTPRequestHandler):
    base_url = ""
    requests: list[tuple[str, str, bytes]] = []
    folders: set[str] = set()
    files: set[str] = set()
    folder_add_attempts: list[str] = []

    def log_message(self, format: str, *args: object) -> None:
        pass

    def send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def record(self) -> tuple[str, bytes]:
        body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
        path = unquote(urlsplit(self.path).path)
        self.requests.append((self.command, self.path, body))
        assert self.headers["Authorization"] == "Bearer test-token"
        return path, body

    def do_GET(self) -> None:
        path, _ = self.record()
        if path == "/sites/team/_api/web/lists":
            self.send_json(
                {
                    "value": [{"Id": "list-1", "Title": "Documents"}],
                    "odata.nextLink": f"{self.base_url}/page/lists-2",
                }
            )
        elif path == "/page/lists-2":
            self.send_json({"value": [{"Id": "list-2", "Title": "Archive"}]})
        elif path.endswith("/items(2)/File"):
            self.send_json(
                {
                    "ServerRelativeUrl": "/sites/team/Documents/report.txt",
                    "ServerRedirectedEmbedUri": (f"{self.base_url}/sites/team/Documents/report.txt?action=interactive"),
                }
            )
        elif "/GetFolderByServerRelativePath(" in path and path.endswith("/Files"):
            server_relative_url = path.split("DecodedUrl='", 1)[1].split("')", 1)[0]
            prefix = f"{server_relative_url}/"
            self.send_json(
                {
                    "value": [
                        {"ServerRelativeUrl": candidate}
                        for candidate in self.files
                        if candidate.startswith(prefix) and "/" not in candidate[len(prefix) :]
                    ]
                }
            )
        elif "/GetFolderByServerRelativePath(" in path and path.endswith("/Folders"):
            server_relative_url = path.split("DecodedUrl='", 1)[1].split("')", 1)[0]
            prefix = f"{server_relative_url}/"
            self.send_json(
                {
                    "value": [
                        {"ServerRelativeUrl": candidate}
                        for candidate in self.folders
                        if candidate.startswith(prefix) and "/" not in candidate[len(prefix) :]
                    ]
                }
            )
        elif "/GetFolderByServerRelativePath(" in path:
            server_relative_url = path.split("DecodedUrl='", 1)[1].rsplit("')", 1)[0]
            if server_relative_url not in self.folders:
                self.send_json({"error": "File Not Found."}, status=404)
            else:
                prefix = f"{server_relative_url}/"
                item_count = sum(
                    candidate.startswith(prefix) and "/" not in candidate[len(prefix) :]
                    for candidate in self.folders | self.files
                )
                self.send_json({"ServerRelativeUrl": server_relative_url, "ItemCount": item_count})
        elif "/GetFileByServerRelativePath(" in path and not path.endswith("/$value"):
            server_relative_url = path.split("DecodedUrl='", 1)[1].rsplit("')", 1)[0]
            self.send_json(
                {
                    "ServerRelativeUrl": server_relative_url,
                    "ServerRedirectedEmbedUri": (f"{self.base_url}{server_relative_url}?action=interactive"),
                    "Name": server_relative_url.rsplit("/", 1)[-1],
                }
            )
        elif "/GetFileById(" in path and not path.endswith("/$value"):
            self.send_json(
                {
                    "ServerRelativeUrl": "/sites/team/Documents/report.txt",
                    "ServerRedirectedEmbedUri": (f"{self.base_url}/sites/team/Documents/report.txt?action=interactive"),
                    "Name": "report.txt",
                }
            )
        elif "/lists/GetByTitle('Documents')/items" in path or "/lists/GetById('list-1')/items" in path:
            self.send_json(
                {
                    "value": [
                        {
                            "Id": 1,
                            "FileSystemObjectType": 1,
                            "ServerRelativeUrl": "/sites/team/Documents/Folder",
                        },
                        {"Id": 2, "FileSystemObjectType": 0, "Title": "report.txt"},
                    ]
                }
            )
        elif path.endswith("/$value"):
            body = b"report-content"
            self.send_response(200)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif path.endswith("/DefaultDocumentLibrary"):
            self.send_json({"Id": "list-1", "Title": "Documents"})
        elif path.endswith("/CurrentUser"):
            self.send_json({"LoginName": "user@example.test"})
        elif path == "/sites/team/_api/site":
            self.send_json({"Id": "site-id"})
        elif path == "/sites/team/_api/web":
            self.send_json({"Id": "web-id"})
        elif path.endswith("/RootFolder"):
            self.send_json({"ServerRelativeUrl": "/sites/team/Documents"})
        elif path == "/sites/team/_api/search/query":
            self.send_json(
                {
                    "PrimaryQueryResult": {
                        "RelevantResults": {
                            "Table": {"Rows": [{"Cells": [{"Key": "Title", "Value": "Quarterly report"}]}]}
                        }
                    }
                }
            )
        elif "/lists/GetByTitle('Documents')" in path or "/lists/GetById('list-1')" in path:
            self.send_json({"Id": "list-1", "Title": "Documents"})
        else:
            self.send_json({"error": path}, status=404)

    def do_POST(self) -> None:
        path, body = self.record()
        if path.endswith("/GetItems"):
            view_xml = json.loads(body)["query"]["ViewXml"]
            if view_xml == "<View />":
                self.send_json({"value": []})
            else:
                self.send_json(
                    {
                        "value": [
                            {
                                "Id": 1,
                                "FileSystemObjectType": 1,
                                "ServerRelativeUrl": "/sites/team/Documents/Folder",
                            },
                            {"Id": 2, "FileSystemObjectType": 0, "Title": "report.txt"},
                        ]
                    }
                )
        elif "/Files/add(" in path:
            folder_path = path.split("GetFolderByServerRelativeUrl('", 1)[1].split("')", 1)[0]
            if folder_path not in self.folders:
                self.send_json({"error": "File Not Found."}, status=404)
                return
            filename = path.split("Files/add(url='", 1)[1].split("'", 1)[0]
            file_path = f"{folder_path}/{filename}"
            self.files.add(file_path)
            self.send_json({"ServerRelativeUrl": file_path})
        elif "/startUpload(" in path or "/continueUpload(" in path:
            self.send_json({})
        elif "/finishUpload(" in path:
            self.send_json({"ServerRelativeUrl": "/sites/team/Documents/large.bin"})
        elif "/Folders/AddUsingPath(" in path:
            assert self.headers["Content-Length"] == "0"
            folder_path = path.split("DecodedUrl='", 1)[1].split("'", 1)[0]
            self.folder_add_attempts.append(folder_path)
            if folder_path in self.folders:
                self.send_json({"error": "Folder already exists."}, status=409)
                return
            parent_path = folder_path.rsplit("/", 1)[0]
            if parent_path not in self.folders:
                self.send_json({"error": "File Not Found."}, status=404)
                return
            self.folders.add(folder_path)
            self.send_json({"ServerRelativeUrl": folder_path})
        elif "/GetUserEffectivePermissions(" in path:
            self.send_json({"High": "16", "Low": "0"})
        else:
            self.send_json({"error": path}, status=404)

    def do_DELETE(self) -> None:
        path, _ = self.record()
        assert self.headers["If-Match"] == "*"
        if "/GetFolderByServerRelativePath(" in path:
            folder_path = path.split("DecodedUrl='", 1)[1].rsplit("')", 1)[0]
            if folder_path not in self.folders:
                self.send_json({"error": "File Not Found."}, status=404)
                return
            prefix = f"{folder_path}/"
            if any(candidate.startswith(prefix) for candidate in self.folders | self.files):
                self.send_json({"error": "Folder is not empty."}, status=500)
                return
            self.folders.remove(folder_path)
            self.send_response(204)
            self.end_headers()
        elif "/GetFileByServerRelativePath(" in path:
            file_path = path.split("DecodedUrl='", 1)[1].rsplit("')", 1)[0]
            if file_path not in self.files:
                self.send_json({"error": "File Not Found."}, status=404)
                return
            self.files.remove(file_path)
            self.send_response(204)
            self.end_headers()
        else:
            self.send_json({"error": path}, status=404)


@contextmanager
def sharepoint_server() -> Iterator[str]:
    server = ThreadingHTTPServer(("127.0.0.1", 0), SharePointHandler)
    SharePointHandler.base_url = f"http://127.0.0.1:{server.server_port}"
    SharePointHandler.requests = []
    SharePointHandler.folders = {
        "/sites/team/Documents",
        "/sites/team/Documents/Folder",
    }
    SharePointHandler.files = {"/sites/team/Documents/report.txt"}
    SharePointHandler.folder_add_attempts = []
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield f"{SharePointHandler.base_url}/sites/team"
    finally:
        server.shutdown()
        thread.join()
        server.server_close()


@pytest.mark.asyncio
async def test_native_sharepoint_operations() -> None:
    with sharepoint_server() as site_url:
        async with SharePointClient.from_static_token(site_url, "test-token") as client:
            initial_lists, remaining = await client.get(max_wait=-1)
            assert [item.title for item in initial_lists] == ["Documents"]
            assert [item.title for item in await remaining] == ["Archive"]

            documents = await client.get("Documents")
            assert documents.title == "Documents"
            assert (await client.get_default_document_library()).id == "list-1"

            items = await documents.get_items()
            assert isinstance(items[0], SPFolder)
            assert items[0].server_relative_url == "/sites/team/Documents/Folder"
            assert isinstance(items[1], SPFile)
            await items[1].resolve()
            assert items[1].properties["ServerRelativeUrl"] == "/sites/team/Documents/report.txt"
            assert await items[1].download() == b"report-content"

            file = await client.get_file("/sites/team/Documents/report.txt")
            assert file.resolved
            assert file.properties["ServerRelativeUrl"] == "/sites/team/Documents/report.txt"
            assert file.properties["Name"] == "report.txt"
            assert "action=default" in file.get_url()
            assert file.browser_url().endswith(
                "Documents/Forms/AllItems.aspx?id=%2Fsites%2Fteam%2FDocuments%2Freport.txt&parent=%2Fsites%2Fteam%2FDocuments"
            )
            assert await file.download() == b"report-content"

            browser_file_url = f"{site_url}/Documents/Forms/AllItems.aspx?id=%2Fsites%2Fteam%2FDocuments%2Freport.txt"
            relative_path = (await client.get_file(browser_file_url)).properties["ServerRelativeUrl"]
            assert relative_path is not None
            assert relative_path.endswith("report.txt")
            assert await client.download(browser_file_url) == b"report-content"

            sourcedoc_url = (
                f"{site_url}/_layouts/15/Doc.aspx?sourcedoc=%7B01246A4B-84D7-49D6-8937-895D3C0F50A9%7D"
                "&file=report.txt&action=edit"
            )
            sourcedoc_file = await client.get_file(sourcedoc_url)
            assert sourcedoc_file.unique_id == "01246A4B-84D7-49D6-8937-895D3C0F50A9"
            assert sourcedoc_file.properties["ServerRelativeUrl"] == "/sites/team/Documents/report.txt"
            assert await client.download(sourcedoc_url) == b"report-content"

            listed = await client.ls("/sites/team/Documents/Folder")
            assert isinstance(listed[0], SPFolder)
            assert isinstance(listed[1], SPFile)
            assert len(await listed[0].ls()) == 2
            assert (await client.get("Documents")).get_url().endswith("Documents/Forms/AllItems.aspx")
            assert listed[0].get_url().endswith("Documents/Forms/AllItems.aspx?id=%2Fsites%2Fteam%2FDocuments%2FFolder")

            caml_items, caml_remaining = await client.get_items("Documents", caml="<View />", max_wait=0)
            assert caml_items == []
            assert await caml_remaining == []

            assert await client.download("/sites/team/Documents/report.txt") == b"report-content"
            uploaded = await client.upload("/sites/team/Documents/upload.txt", b"uploaded content")
            assert uploaded.properties["ServerRelativeUrl"] == "/sites/team/Documents/upload.txt"
            large = await client.upload("/sites/team/Documents/large.bin", b"x" * (8 * 1024 * 1024 + 1))
            assert large.properties["ServerRelativeUrl"] == "/sites/team/Documents/large.bin"
            upload_requests = [
                (path, body)
                for method, path, body in SharePointHandler.requests
                if method == "POST" and "Upload(" in path
            ]
            assert [len(body) for _, body in upload_requests] == [4 * 1024 * 1024, 4 * 1024 * 1024, 1]
            folder = await client.add_folder("/sites/team/Documents/New")
            assert folder.server_relative_url == "/sites/team/Documents/New"
            assert (await client.get_current_user())["LoginName"] == "user@example.test"
            assert (await client.get_effective_permissions("/sites/team/Documents"))["High"] == "16"
            assert (await client.search("quarterly", title="Documents"))[0]["Title"] == "Quarterly report"


@pytest.mark.asyncio
async def test_chunked_and_file_helpers(tmp_path) -> None:
    with sharepoint_server() as site_url:
        async with SharePointClient.from_static_token(site_url, "test-token") as client:
            async with client.upload_chunks("/sites/team/Documents/chunked.txt") as upload:
                await upload.write(b"hello ")
                await upload.write(b"world")

            async with client.download_chunks("/sites/team/Documents/report.txt") as download:
                chunks = []
                while (chunk := await download.get_chunk()) is not None:
                    chunks.append(chunk)
                assert b"".join(chunks) == b"report-content"

            local_download = tmp_path / "downloaded.txt"
            assert await client.download_file("/sites/team/Documents/report.txt", str(local_download)) is None
            assert local_download.read_bytes() == b"report-content"

            file = await client.get_file("/sites/team/Documents/report.txt")
            local_file_download = tmp_path / "downloaded2.txt"
            assert await file.download_file(str(local_file_download)) is None
            assert local_file_download.read_bytes() == b"report-content"

            async with file.download_chunks() as download:
                chunks = []
                while (chunk := await download.get_chunk()) is not None:
                    chunks.append(chunk)
                assert b"".join(chunks) == b"report-content"

            local_upload = tmp_path / "to_upload.bin"
            local_upload.write_bytes(b"y" * (8 * 1024 * 1024 + 1))
            assert await client.upload_file("/sites/team/Documents/upload_from_file.bin", str(local_upload)) is None


@pytest.mark.asyncio
async def test_upload_creates_missing_folders_and_delete_methods(tmp_path) -> None:
    with sharepoint_server() as site_url:
        async with SharePointClient.from_static_token(site_url, "test-token") as client:
            uploaded = await client.upload("/sites/team/Documents/a/b/c/upload.txt", b"content")
            assert uploaded.properties["ServerRelativeUrl"].endswith("/a/b/c/upload.txt")
            assert SharePointHandler.folder_add_attempts == [
                "/sites/team/Documents/a",
                "/sites/team/Documents/a/b",
                "/sites/team/Documents/a/b/c",
            ]
            probed_paths = {
                unquote(urlsplit(request_path).path).split("DecodedUrl='", 1)[1].rsplit("')", 1)[0]
                for method, request_path, _ in SharePointHandler.requests
                if method == "GET" and "/GetFolderByServerRelativePath(" in request_path
            }
            assert {
                "/sites/team/Documents",
                "/sites/team/Documents/a",
                "/sites/team/Documents/a/b",
                "/sites/team/Documents/a/b/c",
            } <= probed_paths
            assert await uploaded.del_file() is None

            local_upload = tmp_path / "local.bin"
            local_upload.write_bytes(b"local")
            assert await client.upload_file("/sites/team/Documents/from/file/helper.bin", str(local_upload)) is None
            assert await client.del_file("/sites/team/Documents/from/file/helper.bin") is None

            async with client.upload_chunks("/sites/team/Documents/chunked/nested/file.bin") as upload:
                await upload.write(b"one")
                await upload.write(b"two")

            folder = await client.add_folder("/sites/team/Documents/object-folder")
            assert await folder.del_folder() is None
            assert await client.del_folder("/sites/team/Documents/object-folder", ignore_missing=True) is None

            await client.upload("/sites/team/Documents/nonempty/child.txt", b"content")
            with pytest.raises(RuntimeError, match="folder is not empty"):
                await client.del_folder("/sites/team/Documents/nonempty")
            assert await client.del_folder("/sites/team/Documents/nonempty", recursive=True) is None
