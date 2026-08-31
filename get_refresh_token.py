"""
get_refresh_token.py — one-off local OAuth flow to get a FreeAgent refresh
token, without your Client Secret passing through any third-party tool.

Prerequisites:
  - You've registered an app at https://dev.freeagent.com/apps and have its
    Client ID and Client Secret.
  - That app has a registered redirect URI matching REDIRECT_URI below
    (default: http://localhost:53682/callback). If you change PORT, update
    the redirect URI in the Developer Dashboard to match exactly.
  - .env has FREEAGENT_CLIENT_ID and FREEAGENT_CLIENT_SECRET filled in.
    FREEAGENT_REFRESH_TOKEN is what this script produces — leave it blank.

Usage:
    python get_refresh_token.py

This starts a local server, opens your browser to FreeAgent's login/approve
screen, catches the redirect, exchanges the authorization code for tokens
(HTTP Basic Auth POST per https://dev.freeagent.com/docs/oauth), and writes
the refresh token straight into .env.

Note: authorization codes expire after 15 minutes, so the exchange happens
immediately after the browser redirect — no separate manual step.
"""
import os
import re
import secrets
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlencode, urlparse, parse_qs

import requests
from dotenv import load_dotenv

load_dotenv()

PORT = 53682
REDIRECT_URI = f"http://localhost:{PORT}/callback"
AUTH_ENDPOINT = "https://api.freeagent.com/v2/approve_app"
TOKEN_ENDPOINT = "https://api.freeagent.com/v2/token_endpoint"
ENV_FILE = ".env"
CALLBACK_TIMEOUT_SECONDS = 300  # 5 minutes to complete the browser flow


class CallbackHandler(BaseHTTPRequestHandler):
    """Captures exactly one redirect to /callback and stores the result on
    the server instance (server.oauth_result) so the caller can read it."""

    def log_message(self, *args):
        pass  # keep the console quiet

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        result = {
            "code": params.get("code", [None])[0],
            "state": params.get("state", [None])[0],
            "error": params.get("error", [None])[0],
        }
        self.server.oauth_result = result

        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        if result["error"]:
            body = f"<html><body><h3>Authorization failed: {result['error']}</h3>You can close this tab.</body></html>"
        else:
            body = "<html><body><h3>Authorized.</h3>You can close this tab and go back to the terminal.</body></html>"
        self.wfile.write(body.encode())


def make_callback_server(port):
    server = HTTPServer(("localhost", port), CallbackHandler)
    server.oauth_result = None
    server.timeout = CALLBACK_TIMEOUT_SECONDS
    return server


def build_auth_url(client_id, redirect_uri, state):
    return AUTH_ENDPOINT + "?" + urlencode({
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": redirect_uri,
        "state": state,
    })


def exchange_code_for_tokens(client_id, client_secret, code, redirect_uri):
    resp = requests.post(
        TOKEN_ENDPOINT,
        auth=(client_id, client_secret),
        data={
            "grant_type": "authorization_code",
            "code": code,
            "redirect_uri": redirect_uri,
        },
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def update_env_refresh_token(refresh_token, env_file=ENV_FILE):
    """Set FREEAGENT_REFRESH_TOKEN in env_file, replacing an existing line
    if present, appending one if not, and leaving every other line alone."""
    if not os.path.exists(env_file):
        print(f"\n{env_file} not found. Add this line to it manually:")
        print(f"FREEAGENT_REFRESH_TOKEN={refresh_token}")
        return False

    with open(env_file) as f:
        content = f.read()

    if re.search(r"^FREEAGENT_REFRESH_TOKEN=.*$", content, flags=re.MULTILINE):
        content = re.sub(
            r"^FREEAGENT_REFRESH_TOKEN=.*$",
            f"FREEAGENT_REFRESH_TOKEN={refresh_token}",
            content,
            flags=re.MULTILINE,
        )
    else:
        content = content.rstrip("\n") + f"\nFREEAGENT_REFRESH_TOKEN={refresh_token}\n"

    with open(env_file, "w") as f:
        f.write(content)
    print(f"\nWrote FREEAGENT_REFRESH_TOKEN into {env_file}.")
    return True


def run_oauth_flow(client_id, client_secret, port=PORT, open_browser=True):
    """Runs the full flow and returns the token response dict. Raises
    SystemExit with a clear message on any failure mode."""
    redirect_uri = f"http://localhost:{port}/callback"
    state = secrets.token_urlsafe(16)
    auth_url = build_auth_url(client_id, redirect_uri, state)

    print("Opening your browser to authorize this app against FreeAgent...")
    print(f"If it doesn't open automatically, visit:\n{auth_url}\n")
    if open_browser:
        webbrowser.open(auth_url)

    server = make_callback_server(port)
    print(f"Waiting for the redirect on {redirect_uri} ({CALLBACK_TIMEOUT_SECONDS}s timeout)...")
    server.handle_request()
    result = server.oauth_result

    if not result or not result.get("code"):
        if result and result.get("error"):
            sys.exit(f"FreeAgent returned an error: {result['error']}")
        sys.exit("Timed out waiting for authorization. Re-run and try again.")

    if result.get("state") != state:
        sys.exit("State mismatch on the callback — possible interference. Re-run to try again.")

    tokens = exchange_code_for_tokens(client_id, client_secret, result["code"], redirect_uri)
    return tokens


def main():
    client_id = os.environ.get("FREEAGENT_CLIENT_ID")
    client_secret = os.environ.get("FREEAGENT_CLIENT_SECRET")
    if not client_id or not client_secret:
        sys.exit(
            "Set FREEAGENT_CLIENT_ID and FREEAGENT_CLIENT_SECRET in .env first "
            "(from https://dev.freeagent.com/apps), then re-run this."
        )

    tokens = run_oauth_flow(client_id, client_secret)

    print(
        "\nSuccess. Access token (expires in "
        f"{tokens.get('expires_in', '?')}s) and refresh token obtained."
    )
    update_env_refresh_token(tokens["refresh_token"])
    print(
        "You shouldn't need to run this again — freeagent_client.py refreshes "
        "the access token from the stored refresh token automatically."
    )


if __name__ == "__main__":
    main()
