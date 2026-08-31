import threading
from unittest.mock import patch

import pytest
import requests

import get_refresh_token as grt


class TestUpdateEnvRefreshToken:
    def test_replaces_existing_line_without_touching_others(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text(
            "FREEAGENT_CLIENT_ID=abc\n"
            "FREEAGENT_CLIENT_SECRET=def\n"
            "FREEAGENT_REFRESH_TOKEN=old-token\n"
            "ACCOUNT_POLLUTED_ID=1\n"
        )
        grt.update_env_refresh_token("new-token", env_file=str(env_file))

        content = env_file.read_text()
        assert "FREEAGENT_REFRESH_TOKEN=new-token" in content
        assert "old-token" not in content
        assert "FREEAGENT_CLIENT_ID=abc" in content
        assert "ACCOUNT_POLLUTED_ID=1" in content
        # exactly one refresh token line, not duplicated
        assert content.count("FREEAGENT_REFRESH_TOKEN=") == 1

    def test_appends_when_line_missing(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text("FREEAGENT_CLIENT_ID=abc\nFREEAGENT_CLIENT_SECRET=def\n")
        grt.update_env_refresh_token("brand-new-token", env_file=str(env_file))

        content = env_file.read_text()
        assert "FREEAGENT_REFRESH_TOKEN=brand-new-token" in content
        assert "FREEAGENT_CLIENT_ID=abc" in content

    def test_appends_cleanly_when_file_has_no_trailing_newline(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text("FREEAGENT_CLIENT_ID=abc")  # no trailing \n
        grt.update_env_refresh_token("tok", env_file=str(env_file))

        lines = env_file.read_text().splitlines()
        assert lines == ["FREEAGENT_CLIENT_ID=abc", "FREEAGENT_REFRESH_TOKEN=tok"]

    def test_missing_file_prints_instructions_instead_of_crashing(self, tmp_path, capsys):
        missing_path = str(tmp_path / "does_not_exist.env")
        result = grt.update_env_refresh_token("tok", env_file=missing_path)
        assert result is False
        out = capsys.readouterr().out
        assert "FREEAGENT_REFRESH_TOKEN=tok" in out

    def test_blank_existing_value_is_replaced(self, tmp_path):
        env_file = tmp_path / ".env"
        env_file.write_text("FREEAGENT_REFRESH_TOKEN=\n")
        grt.update_env_refresh_token("filled-in", env_file=str(env_file))
        assert "FREEAGENT_REFRESH_TOKEN=filled-in" in env_file.read_text()


class TestBuildAuthUrl:
    def test_includes_required_params_url_escaped(self):
        url = grt.build_auth_url("my-client-id", "http://localhost:53682/callback", "xyz-state")
        assert url.startswith(grt.AUTH_ENDPOINT + "?")
        assert "client_id=my-client-id" in url
        assert "response_type=code" in url
        assert "state=xyz-state" in url
        # redirect_uri must be URL-escaped in the query string
        assert "redirect_uri=http%3A%2F%2Flocalhost%3A53682%2Fcallback" in url


class TestCallbackServerRoundTrip:
    """Runs an actual HTTPServer on a free local port and hits it with a
    real HTTP request, the same way FreeAgent's redirect would."""

    def _serve_one_request_in_background(self, server):
        thread = threading.Thread(target=server.handle_request)
        thread.start()
        return thread

    def test_captures_code_and_state_from_real_request(self):
        server = grt.make_callback_server(0)  # port 0 = OS picks a free port
        port = server.server_port
        thread = self._serve_one_request_in_background(server)

        resp = requests.get(
            f"http://localhost:{port}/callback",
            params={"code": "auth-code-123", "state": "state-abc"},
            timeout=5,
        )
        thread.join(timeout=5)

        assert resp.status_code == 200
        assert "Authorized" in resp.text
        assert server.oauth_result == {"code": "auth-code-123", "state": "state-abc", "error": None}

    def test_captures_error_param_when_user_denies_access(self):
        server = grt.make_callback_server(0)
        port = server.server_port
        thread = self._serve_one_request_in_background(server)

        resp = requests.get(
            f"http://localhost:{port}/callback",
            params={"error": "access_denied"},
            timeout=5,
        )
        thread.join(timeout=5)

        assert "failed" in resp.text.lower()
        assert server.oauth_result["error"] == "access_denied"
        assert server.oauth_result["code"] is None

    def test_times_out_cleanly_when_nothing_arrives(self):
        server = grt.make_callback_server(0)
        server.timeout = 0.3  # don't actually wait 5 minutes in a test
        server.handle_request()
        assert server.oauth_result is None


class TestRunOauthFlowEndToEnd:
    """Simulates the whole flow: a background 'browser' hits the real local
    callback server, and the token exchange itself is mocked (no live
    FreeAgent call)."""

    def test_full_flow_writes_expected_tokens(self, monkeypatch):
        # Pin the port so we can predict the redirect_uri the "browser" hits,
        # and skip actually opening a browser window.
        test_port = 53699

        def fake_browser_hits_callback(auth_url):
            # Extract the state FreeAgent would echo back, then simulate the
            # redirect a real browser would perform after the user approves.
            from urllib.parse import urlparse, parse_qs
            state = parse_qs(urlparse(auth_url).query)["state"][0]

            def hit_callback():
                requests.get(
                    f"http://localhost:{test_port}/callback",
                    params={"code": "the-auth-code", "state": state},
                    timeout=5,
                )

            threading.Timer(0.1, hit_callback).start()

        monkeypatch.setattr("webbrowser.open", fake_browser_hits_callback)

        with patch("get_refresh_token.exchange_code_for_tokens") as mock_exchange:
            mock_exchange.return_value = {
                "access_token": "fake-access",
                "refresh_token": "fake-refresh",
                "expires_in": 3600,
            }
            tokens = grt.run_oauth_flow("client-id", "client-secret", port=test_port)

        assert tokens["refresh_token"] == "fake-refresh"
        mock_exchange.assert_called_once_with(
            "client-id", "client-secret", "the-auth-code", f"http://localhost:{test_port}/callback"
        )

    def test_state_mismatch_aborts_before_token_exchange(self, monkeypatch):
        test_port = 53698

        def fake_browser_hits_callback_with_wrong_state(auth_url):
            def hit_callback():
                requests.get(
                    f"http://localhost:{test_port}/callback",
                    params={"code": "the-auth-code", "state": "WRONG-STATE"},
                    timeout=5,
                )
            threading.Timer(0.1, hit_callback).start()

        monkeypatch.setattr("webbrowser.open", fake_browser_hits_callback_with_wrong_state)

        with patch("get_refresh_token.exchange_code_for_tokens") as mock_exchange:
            with pytest.raises(SystemExit, match="State mismatch"):
                grt.run_oauth_flow("client-id", "client-secret", port=test_port)
        mock_exchange.assert_not_called()

    def test_access_denied_aborts_before_token_exchange(self, monkeypatch):
        test_port = 53697

        def fake_browser_denies(auth_url):
            def hit_callback():
                requests.get(
                    f"http://localhost:{test_port}/callback",
                    params={"error": "access_denied"},
                    timeout=5,
                )
            threading.Timer(0.1, hit_callback).start()

        monkeypatch.setattr("webbrowser.open", fake_browser_denies)

        with patch("get_refresh_token.exchange_code_for_tokens") as mock_exchange:
            with pytest.raises(SystemExit, match="access_denied"):
                grt.run_oauth_flow("client-id", "client-secret", port=test_port)
        mock_exchange.assert_not_called()
