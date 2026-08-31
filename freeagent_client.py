"""
freeagent_client.py — shared helpers for talking to the FreeAgent API.

Auth model: OAuth 2.0, refresh-token grant. You do the interactive
authorization once (see README.md) to obtain a refresh token and store it
in .env. From then on this module mints/refreshes access tokens for you
automatically — nothing here does the initial authorization-code exchange.

Reference:
    https://dev.freeagent.com/docs/oauth
    https://dev.freeagent.com/docs/bank_accounts
    https://dev.freeagent.com/docs/bank_transactions
    https://dev.freeagent.com/docs/bank_transaction_explanations
    https://dev.freeagent.com/docs/introduction  (pagination, rate limits)
"""
import os
import time

import requests
from dotenv import load_dotenv
from get_refresh_token import update_env_refresh_token

load_dotenv()

API_BASE = "https://api.freeagent.com/v2"
TOKEN_ENDPOINT = "https://api.freeagent.com/v2/token_endpoint"

CLIENT_ID = os.environ["FREEAGENT_CLIENT_ID"]
CLIENT_SECRET = os.environ["FREEAGENT_CLIENT_SECRET"]
REFRESH_TOKEN = os.environ["FREEAGENT_REFRESH_TOKEN"]

# Documented limits (dev.freeagent.com/docs/introduction): 120 requests/min,
# 3600/hour, 15 token refreshes/min. We stay well under these with a small
# sleep between paginated calls rather than trying to be clever about it.
_PAGE_SLEEP_SECONDS = 0.3

_access_token = None
_access_token_expiry = 0.0  # unix timestamp


class AuthenticationError(RuntimeError):
    """Raised when FreeAgent refuses an access-token refresh."""


def _refresh_access_token():
    global _access_token, _access_token_expiry, REFRESH_TOKEN
    resp = requests.post(
        TOKEN_ENDPOINT,
        auth=(CLIENT_ID, CLIENT_SECRET),
        data={"grant_type": "refresh_token", "refresh_token": REFRESH_TOKEN},
        timeout=30,
    )
    if resp.status_code >= 400:
        # Do not let callers retry the same invalid refresh credentials once
        # per workbook row. In particular, repeated 401s quickly turn into
        # 429s at the token endpoint.
        raise AuthenticationError(
            f"FreeAgent token refresh failed (HTTP {resp.status_code})"
        )
    resp.raise_for_status()
    data = resp.json()
    _access_token = data["access_token"]
    # Refresh a bit early rather than cutting it exactly at expiry.
    _access_token_expiry = time.time() + data.get("expires_in", 3600) - 300

    new_refresh = data.get("refresh_token")
    if new_refresh and new_refresh != REFRESH_TOKEN:
        REFRESH_TOKEN = new_refresh
        update_env_refresh_token(new_refresh)
        print("FreeAgent issued a new refresh token; now saved to .env.")


def _access_headers():
    if _access_token is None or time.time() >= _access_token_expiry:
        _refresh_access_token()
    return {"Authorization": f"Bearer {_access_token}", "Accept": "application/json"}


def _request(method, path_or_url, **kwargs):
    url = path_or_url if path_or_url.startswith("http") else f"{API_BASE}{path_or_url}"
    headers = kwargs.pop("headers", {})
    headers.update(_access_headers())
    for attempt in range(5):
        resp = requests.request(method, url, headers=headers, timeout=30, **kwargs)
        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", 5))
            print(f"  Rate limited, waiting {wait}s...")
            time.sleep(wait)
            continue
        return resp
    return resp  # last attempt's response, even if still a 429


def get_bank_accounts():
    resp = _request("GET", "/bank_accounts")
    resp.raise_for_status()
    return resp.json()["bank_accounts"]


def get_bank_account(account_id):
    resp = _request("GET", f"/bank_accounts/{account_id}")
    resp.raise_for_status()
    return resp.json()["bank_account"]


def get_bank_transactions(bank_account_url, view="all"):
    """Fetch ALL bank transactions for an account, following pagination.

    Pagination is page/per_page (max 100/page) per
    https://dev.freeagent.com/docs/introduction — we just page until a
    short batch tells us we've hit the end, rather than parsing the Link
    header, which is simpler and equally reliable here.
    """
    transactions = []
    page = 1
    per_page = 100
    while True:
        resp = _request(
            "GET",
            "/bank_transactions",
            params={
                "bank_account": bank_account_url,
                "view": view,
                "page": page,
                "per_page": per_page,
            },
        )
        resp.raise_for_status()
        batch = resp.json()["bank_transactions"]
        transactions.extend(batch)
        if len(batch) < per_page:
            break
        page += 1
        time.sleep(_PAGE_SLEEP_SECONDS)
    return transactions


def get_transaction(transaction_url):
    resp = _request("GET", transaction_url)
    resp.raise_for_status()
    return resp.json()["bank_transaction"]


def get_explanation(explanation_url):
    resp = _request("GET", explanation_url)
    resp.raise_for_status()
    return resp.json()["bank_transaction_explanation"]


def delete_explanation(explanation_url):
    resp = _request("DELETE", explanation_url)
    resp.raise_for_status()


def delete_transaction(transaction_url):
    # Transaction URLs returned by FreeAgent use the plural resource path,
    # and DELETE must use that same path. The API documentation currently
    # shows a singular path in this section, but that path returns 404 for
    # valid transactions; preserving the URL's resource path avoids that
    # mismatch.
    url = transaction_url
    resp = _request("DELETE", url)
    resp.raise_for_status()
