"""Tests for `async_sharepoint` package."""

import asyncio
import threading

import pytest

import async_sharepoint
from async_sharepoint.main import SharePointClient


def test_import():
    """Verify the package can be imported."""
    assert async_sharepoint


@pytest.mark.asyncio
async def test_stalled_refresh_does_not_block_later_requests_forever(monkeypatch):
    client = SharePointClient("https://example.sharepoint.com/sites/test", lambda: {})
    client._reset_loop_state()
    await client._token_lock.acquire()
    monkeypatch.setattr("async_sharepoint.main.TOKEN_REFRESH_TIMEOUT", 0.01)

    with pytest.raises(TimeoutError, match="token refresh lock"):
        await asyncio.wait_for(client._headers(), timeout=0.1)

    client._token_lock.release()
    await client.aclose()


@pytest.mark.asyncio
async def test_stalled_token_provider_times_out_without_holding_lock(monkeypatch):
    release_provider = threading.Event()

    def get_token():
        release_provider.wait()
        return {"access_token": "token", "expires_in": 3600}

    client = SharePointClient("https://example.sharepoint.com/sites/test", get_token)
    monkeypatch.setattr("async_sharepoint.main.TOKEN_REFRESH_TIMEOUT", 0.01)

    try:
        with pytest.raises(TimeoutError, match="acquiring a SharePoint access token"):
            await client._headers()
        assert not client._token_lock.locked()
    finally:
        release_provider.set()
        await client.aclose()


@pytest.mark.asyncio
async def test_concurrent_refreshes_share_one_token_acquisition():
    calls = 0

    def get_token():
        nonlocal calls
        calls += 1
        return {"access_token": "token", "expires_in": 3600}

    client = SharePointClient("https://example.sharepoint.com/sites/test", get_token)
    try:
        headers = await asyncio.gather(client._headers(), client._headers())
        assert calls == 1
        assert all(header["Authorization"] == "Bearer token" for header in headers)
    finally:
        await client.aclose()
