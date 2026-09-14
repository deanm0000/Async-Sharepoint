import asyncio
import inspect
from typing import Any, cast

import pytest

import async_sharepoint
from async_sharepoint import SharePointClient, SPFile, SPFolder


def get_token() -> dict:
    return {"access_token": "test-token", "expires_in": 3600}


def test_public_exports() -> None:
    assert async_sharepoint.__all__ == ["SPFile", "SPFolder", "SharePointClient"]
    assert async_sharepoint.SPFile is SPFile
    assert async_sharepoint.SPFolder is SPFolder
    assert async_sharepoint.SharePointClient is SharePointClient


def signature_parameters(callable_object) -> list[tuple[str, inspect._ParameterKind, object]]:
    parameters = [
        (parameter.name, parameter.kind, parameter.default)
        for parameter in inspect.signature(callable_object).parameters.values()
    ]
    if parameters and parameters[0][0] == "self":
        parameters[0] = ("self", inspect.Parameter.POSITIONAL_OR_KEYWORD, parameters[0][2])
    return parameters


def test_public_signatures() -> None:
    positional = inspect.Parameter.POSITIONAL_OR_KEYWORD
    keyword = inspect.Parameter.KEYWORD_ONLY
    required = inspect.Parameter.empty

    assert signature_parameters(SharePointClient) == [
        ("site_url", positional, required),
        ("get_token", positional, required),
        ("properties", keyword, None),
        ("item_id", keyword, None),
        ("list_url", keyword, None),
    ]
    assert signature_parameters(SharePointClient.get) == [
        ("self", positional, required),
        ("title", positional, None),
        ("id", keyword, None),
        ("max_wait", keyword, None),
    ]
    assert signature_parameters(SharePointClient.get_items) == [
        ("self", positional, required),
        ("title", positional, None),
        ("id", keyword, None),
        ("caml", keyword, None),
        ("max_wait", keyword, None),
    ]
    assert signature_parameters(SharePointClient.upload)[-1] == ("overwrite", keyword, True)
    assert signature_parameters(SharePointClient.add_folder)[-1] == ("overwrite", keyword, False)
    assert signature_parameters(SharePointClient.get_effective_permissions)[-1] == ("login_name", keyword, None)
    assert signature_parameters(SharePointClient.search)[-3:] == [
        ("title", keyword, None),
        ("id", keyword, None),
        ("row_limit", keyword, 6),
    ]


@pytest.mark.asyncio
async def test_context_manager() -> None:
    async with SharePointClient("https://example.test/sites/team/", get_token) as client:
        assert client.site_url == "https://example.test/sites/team"


@pytest.mark.asyncio
async def test_argument_validation() -> None:
    client = SharePointClient("https://example.test/sites/team", get_token)
    invalid_get = cast(Any, client.get)
    try:
        with pytest.raises(ValueError, match="specify only one"):
            await invalid_get("Documents", id="list-id")
        with pytest.raises(ValueError, match="max_wait"):
            await invalid_get("Documents", max_wait=1)
        with pytest.raises(ValueError, match="must specify title or id"):
            await client.get_items()
        with pytest.raises(ValueError, match="must specify title or id"):
            await client.search("report")
    finally:
        await client.aclose()


def test_property_fallback_and_file_url() -> None:
    folder = SPFolder(properties={"Name": "Reports"})
    assert folder.Name == "Reports"
    with pytest.raises(AttributeError, match="Missing"):
        _ = folder.Missing

    file = SPFile(
        client=SharePointClient("https://example.test/sites/team", get_token),
        server_relative_path="/sites/team/My File.txt",
        properties={"ServerRedirectedEmbedUri": "https://example.test/sites/team/My File.txt?action=interactive&x=1"},
    )
    try:
        assert file.get_url() == "https://example.test/sites/team/My%20File.txt?action=default&x=1"
    finally:

        async def close_client() -> None:
            await file.client.aclose()

        asyncio.run(close_client())
