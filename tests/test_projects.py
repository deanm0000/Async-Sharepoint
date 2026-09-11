import os
from functools import lru_cache
from pathlib import Path

import msal
import pytest
from dotenv import load_dotenv

from async_sharepoint import SharePointClient

load_dotenv()


@lru_cache(maxsize=1)
def _get_private_key(cert_path: str, all: bool = False):
    with Path(cert_path).open("r", encoding="utf-8") as f:
        private_key_pem = f.read()
    if all:
        return private_key_pem
    lines = private_key_pem.splitlines()
    private_lines = []
    found_private_key = False
    for line in lines:
        if "BEGIN PRIVATE KEY" in line:
            found_private_key = True
            private_lines.append(line)
        elif "END PRIVATE KEY" in line:
            private_lines.append(line)
            break
        elif found_private_key:
            private_lines.append(line)
    return "\n".join(private_lines)


@lru_cache(maxsize=1)
def get_token_dev(scope=os.environ["TEST_SCOPE"]):
    cert_path = os.environ["TEST_LOCAL_CERT_PATH"]
    tenant_id = os.environ["TEST_TENANT_ID"]
    client_id = os.environ["TEST_CLIENT_ID"]
    cert_settings = {
        "tenant": os.environ["TEST_TENANT"],
        "client_id": client_id,
        "thumbprint": os.environ["TEST_THUMBPRINT"],
        "cert_path": cert_path,
    }

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
        client_id=cert_settings["client_id"],
        authority=authority,
        client_credential={
            "thumbprint": cert_settings["thumbprint"],
            "private_key": _get_private_key(cert_path, True),
        },
    )

    token_result = app.acquire_token_for_client(scopes=[scope])
    assert token_result is not None
    return token_result


@pytest.mark.asyncio
async def test_main():
    ctx = SharePointClient(os.environ["TEST_SITE"], get_token_dev)

    items = await ctx.get_items("Projects")
    assert len(items) > 0
    ctx2 = [x for x in items if isinstance(x, SharePointClient)]
    if ctx2:
        ctx2 = ctx2[0]
    assert isinstance(ctx2, SharePointClient)

    items2, task2 = await ctx2.get_items("Documents", max_wait=2)
