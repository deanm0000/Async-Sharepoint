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
            self.send_json({"ServerRelativeUrl": "/sites/team/Documents/upload.txt"})
        elif "/startUpload(" in path or "/continueUpload(" in path:
            self.send_json({})
        elif "/finishUpload(" in path:
            self.send_json({"ServerRelativeUrl": "/sites/team/Documents/large.bin"})
        elif "/Folders/AddUsingPath(" in path:
            self.send_json({"ServerRelativeUrl": "/sites/team/Documents/New"})
        elif "/GetUserEffectivePermissions(" in path:
            self.send_json({"High": "16", "Low": "0"})
        else:
            self.send_json({"error": path}, status=404)


@contextmanager
def sharepoint_server() -> Iterator[str]:
    server = ThreadingHTTPServer(("127.0.0.1", 0), SharePointHandler)
    SharePointHandler.base_url = f"http://127.0.0.1:{server.server_port}"
    SharePointHandler.requests = []
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
            assert items[1].server_relative_path == "/sites/team/Documents/report.txt"
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
            relative_path = (await client.get_file(browser_file_url)).server_relative_path
            assert relative_path is not None
            assert relative_path.endswith("report.txt")
            assert await client.download(browser_file_url) == b"report-content"

            sourcedoc_url = (
                f"{site_url}/_layouts/15/Doc.aspx?sourcedoc=%7B01246A4B-84D7-49D6-8937-895D3C0F50A9%7D"
                "&file=report.txt&action=edit"
            )
            sourcedoc_file = await client.get_file(sourcedoc_url)
            assert sourcedoc_file.unique_id == "01246A4B-84D7-49D6-8937-895D3C0F50A9"
            assert sourcedoc_file.server_relative_path == "/sites/team/Documents/report.txt"
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
            uploaded = await client.upload("/sites/team/Documents", "upload.txt", b"uploaded content")
            assert uploaded.server_relative_path == "/sites/team/Documents/upload.txt"
            large = await client.upload("/sites/team/Documents", "large.bin", b"x" * (8 * 1024 * 1024 + 1))
            assert large.server_relative_path == "/sites/team/Documents/large.bin"
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
