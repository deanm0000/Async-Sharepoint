import os

import pytest
from dotenv import load_dotenv

from async_sharepoint import CertificateCredential, SharePointClient

load_dotenv()


def get_credential_dev() -> CertificateCredential:
    return CertificateCredential(
        tenant_id=os.environ["TEST_TENANT_ID"],
        client_id=os.environ["TEST_CLIENT_ID"],
        private_key_path=os.environ["TEST_LOCAL_CERT_PATH"],
        thumbprint=os.environ["TEST_THUMBPRINT"],
    )


@pytest.mark.asyncio
async def test_main():
    ctx = SharePointClient(os.environ["TEST_SITE"], get_credential_dev())

    items = await ctx.get_items("Projects")
    assert len(items) > 0
    ctx2 = [x for x in items if isinstance(x, SharePointClient)]
    if ctx2:
        ctx2 = ctx2[0]
    assert isinstance(ctx2, SharePointClient)

    items2, task2 = await ctx2.get_items("Documents", max_wait=2)
