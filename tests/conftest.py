import os
import sys

# freeagent_client.py reads these from the environment at import time.
# Set dummies before anything imports it, so importing the module under
# test never depends on a real .env file being present.
os.environ.setdefault("FREEAGENT_CLIENT_ID", "test-client-id")
os.environ.setdefault("FREEAGENT_CLIENT_SECRET", "test-client-secret")
os.environ.setdefault("FREEAGENT_REFRESH_TOKEN", "test-refresh-token")

# The scripts under test live one directory up from tests/, with no
# package structure, so put that directory on sys.path.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest  # noqa: E402
import freeagent_client as fa  # noqa: E402


@pytest.fixture(autouse=True)
def reset_fa_auth_state():
    """freeagent_client caches the access token at module level. Reset that
    cache before and after every test so tests can't leak state into each
    other regardless of run order."""
    fa._access_token = None
    fa._access_token_expiry = 0.0
    fa.REFRESH_TOKEN = "test-refresh-token"
    yield
    fa._access_token = None
    fa._access_token_expiry = 0.0
    fa.REFRESH_TOKEN = "test-refresh-token"
