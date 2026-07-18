"""
GitHub App authentication for the "Refresh Secondary Trades" button.

Mints a short-lived installation access token on demand from the
bond-tracker-refresh GitHub App, so the Streamlit app can dispatch the
secondary-trades-refresh workflow without storing a long-lived PAT.

Only the App's private key is secret; it lives in Streamlit secrets as
GH_APP_PRIVATE_KEY. The App ID and Installation ID are not sensitive (the
repo is public) and are hard-coded below. This replaces the old
GH_WORKFLOW_TOKEN personal-access-token approach.
"""

import time

import jwt  # PyJWT — needs the `cryptography` package for RS256
import requests
import streamlit as st

# Not secret — safe to hard-code.
APP_ID = "4326382"                       # numeric GitHub App ID
INSTALLATION_ID = "147266129"            # installation on k13-spec/bond-tracker

_API = "https://api.github.com"


def get_installation_token() -> str:
    """Return a fresh (~1 hour) installation access token, or raise on failure."""
    private_key = st.secrets["GH_APP_PRIVATE_KEY"]  # PEM string

    # App JWT: valid for < 10 min. iat backdated 60s to tolerate clock skew.
    now = int(time.time())
    payload = {"iat": now - 60, "exp": now + 540, "iss": str(APP_ID)}
    app_jwt = jwt.encode(payload, private_key, algorithm="RS256")

    resp = requests.post(
        f"{_API}/app/installations/{INSTALLATION_ID}/access_tokens",
        headers={
            "Authorization": f"Bearer {app_jwt}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        },
        timeout=20,
    )
    if resp.status_code != 201:
        raise RuntimeError(
            f"installation token request failed ({resp.status_code}): "
            f"{resp.text[:200]}"
        )
    return resp.json()["token"]
