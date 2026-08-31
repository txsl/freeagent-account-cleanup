import time
from unittest.mock import Mock, patch

import freeagent_client as fa


def _mock_response(json_data, status_code=200, headers=None):
    resp = Mock()
    resp.status_code = status_code
    resp.headers = headers or {}
    resp.json.return_value = json_data
    resp.raise_for_status = Mock()
    if status_code >= 400:
        resp.raise_for_status.side_effect = Exception(f"HTTP {status_code}")
    return resp


class TestTokenRefresh:
    def test_refresh_sets_access_token_and_expiry_with_buffer(self):
        with patch("freeagent_client.requests.post") as mock_post:
            mock_post.return_value = _mock_response({
                "access_token": "new-access-token",
                "expires_in": 3600,
                "refresh_token": "test-refresh-token",  # unchanged
            })
            fa._refresh_access_token()

        assert fa._access_token == "new-access-token"
        # Should refresh ~5 min before actual expiry, not exactly at it.
        assert fa._access_token_expiry > time.time() + 3000
        assert fa._access_token_expiry < time.time() + 3300

    def test_refresh_sends_credentials_and_grant_type(self):
        with patch("freeagent_client.requests.post") as mock_post:
            mock_post.return_value = _mock_response({"access_token": "x", "expires_in": 3600})
            fa._refresh_access_token()

        _, kwargs = mock_post.call_args
        assert kwargs["auth"] == ("test-client-id", "test-client-secret")
        assert kwargs["data"]["grant_type"] == "refresh_token"
        assert kwargs["data"]["refresh_token"] == "test-refresh-token"

    def test_refresh_updates_stored_token_when_rotated(self, capsys):
        with patch("freeagent_client.requests.post") as mock_post:
            mock_post.return_value = _mock_response({
                "access_token": "tok",
                "expires_in": 3600,
                "refresh_token": "brand-new-refresh-token",
            })
            fa._refresh_access_token()

        assert fa.REFRESH_TOKEN == "brand-new-refresh-token"
        assert "brand-new-refresh-token" in capsys.readouterr().out

    def test_refresh_leaves_stored_token_unchanged_when_not_rotated(self):
        with patch("freeagent_client.requests.post") as mock_post:
            mock_post.return_value = _mock_response({
                "access_token": "tok",
                "expires_in": 3600,
                "refresh_token": "test-refresh-token",
            })
            fa._refresh_access_token()

        assert fa.REFRESH_TOKEN == "test-refresh-token"


class TestAccessHeaders:
    def test_missing_token_triggers_refresh(self):
        def fake_refresh():
            fa._access_token = "abc"
            fa._access_token_expiry = time.time() + 3600

        with patch("freeagent_client._refresh_access_token", side_effect=fake_refresh) as mock_refresh:
            headers = fa._access_headers()

        mock_refresh.assert_called_once()
        assert headers["Authorization"] == "Bearer abc"

    def test_expired_token_triggers_refresh(self):
        fa._access_token = "stale"
        fa._access_token_expiry = time.time() - 1

        def fake_refresh():
            fa._access_token = "fresh"
            fa._access_token_expiry = time.time() + 3600

        with patch("freeagent_client._refresh_access_token", side_effect=fake_refresh) as mock_refresh:
            headers = fa._access_headers()

        mock_refresh.assert_called_once()
        assert headers["Authorization"] == "Bearer fresh"

    def test_valid_token_is_reused_without_refresh(self):
        fa._access_token = "cached"
        fa._access_token_expiry = time.time() + 1000

        with patch("freeagent_client._refresh_access_token") as mock_refresh:
            headers = fa._access_headers()

        mock_refresh.assert_not_called()
        assert headers["Authorization"] == "Bearer cached"


class TestRequestRetry:
    def test_retries_on_429_then_returns_success(self, monkeypatch):
        fa._access_token = "tok"
        fa._access_token_expiry = time.time() + 3600
        monkeypatch.setattr(fa.time, "sleep", lambda s: None)

        responses = [
            _mock_response({}, status_code=429, headers={"Retry-After": "0"}),
            _mock_response({"ok": True}, status_code=200),
        ]
        with patch("freeagent_client.requests.request", side_effect=responses) as mock_req:
            resp = fa._request("GET", "/company")

        assert resp.status_code == 200
        assert mock_req.call_count == 2

    def test_gives_up_after_five_attempts_of_429(self, monkeypatch):
        fa._access_token = "tok"
        fa._access_token_expiry = time.time() + 3600
        monkeypatch.setattr(fa.time, "sleep", lambda s: None)

        with patch(
            "freeagent_client.requests.request",
            return_value=_mock_response({}, status_code=429, headers={}),
        ) as mock_req:
            resp = fa._request("GET", "/company")

        assert resp.status_code == 429
        assert mock_req.call_count == 5


class TestApiWrappers:
    def test_get_bank_accounts_unwraps_list(self):
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({"bank_accounts": [{"url": "a"}, {"url": "b"}]})
            accounts = fa.get_bank_accounts()
        assert accounts == [{"url": "a"}, {"url": "b"}]

    def test_get_bank_account_unwraps_singular_key_and_requests_correct_path(self):
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({"bank_account": {"current_balance": "1234.56"}})
            account = fa.get_bank_account("7")
        mock_req.assert_called_once_with("GET", "/bank_accounts/7")
        assert account == {"current_balance": "1234.56"}

    def test_get_bank_transactions_paginates_until_short_page(self, monkeypatch):
        monkeypatch.setattr(fa.time, "sleep", lambda s: None)
        page1 = [{"url": f"t{i}"} for i in range(100)]
        page2 = [{"url": f"t{i}"} for i in range(100, 200)]
        page3 = [{"url": f"t{i}"} for i in range(200, 230)]  # short page ends pagination

        with patch(
            "freeagent_client._request",
            side_effect=[
                _mock_response({"bank_transactions": page1}),
                _mock_response({"bank_transactions": page2}),
                _mock_response({"bank_transactions": page3}),
            ],
        ) as mock_req:
            result = fa.get_bank_transactions("https://api.freeagent.com/v2/bank_accounts/1")

        assert len(result) == 230
        assert mock_req.call_count == 3
        pages_requested = [c.kwargs["params"]["page"] for c in mock_req.call_args_list]
        assert pages_requested == [1, 2, 3]

    def test_get_bank_transactions_handles_exact_multiple_of_page_size(self, monkeypatch):
        # If the account has exactly 100 transactions, a follow-up call
        # returning [] is needed to know there isn't a page 2.
        monkeypatch.setattr(fa.time, "sleep", lambda s: None)
        page1 = [{"url": f"t{i}"} for i in range(100)]

        with patch(
            "freeagent_client._request",
            side_effect=[
                _mock_response({"bank_transactions": page1}),
                _mock_response({"bank_transactions": []}),
            ],
        ) as mock_req:
            result = fa.get_bank_transactions("https://api.freeagent.com/v2/bank_accounts/1")

        assert len(result) == 100
        assert mock_req.call_count == 2

    def test_get_transaction_unwraps_singular_key(self):
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({"bank_transaction": {"url": "x"}})
            txn = fa.get_transaction("https://api.freeagent.com/v2/bank_transactions/5")
        assert txn == {"url": "x"}

    def test_get_explanation_unwraps_singular_key(self):
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({"bank_transaction_explanation": {"is_deletable": True}})
            exp = fa.get_explanation("https://api.freeagent.com/v2/bank_transaction_explanations/9")
        assert exp == {"is_deletable": True}

    def test_delete_explanation_issues_delete_to_given_url(self):
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({})
            fa.delete_explanation("https://api.freeagent.com/v2/bank_transaction_explanations/9")
        mock_req.assert_called_once_with(
            "DELETE", "https://api.freeagent.com/v2/bank_transaction_explanations/9"
        )

    def test_delete_transaction_uses_documented_singular_endpoint(self):
        # Regression guard: FreeAgent's docs show this against
        # /v2/bank_transaction/:id (singular), unlike every other endpoint.
        # If this ever needs to change, it should be a deliberate edit here,
        # not an accidental one.
        with patch("freeagent_client._request") as mock_req:
            mock_req.return_value = _mock_response({})
            fa.delete_transaction("https://api.freeagent.com/v2/bank_transactions/42")
        mock_req.assert_called_once_with(
            "DELETE", "https://api.freeagent.com/v2/bank_transaction/42"
        )
