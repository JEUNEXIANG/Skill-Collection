"""
gsheets_util.py — Shared Google Sheets auth + helpers for Platform SOB Agent.
All API calls in this directory should import from here.
Handles:
  - macOS SSL cert fix (certifi)
  - Token refresh when expired
  - Token save-back
"""
import json, os, urllib.request, urllib.parse, ssl, datetime
import certifi

TOKEN_PATH = os.path.expanduser("~/.hermes/google_token.json")
SSL_CTX = ssl.create_default_context(cafile=certifi.where())

SHEETS = {
    "Weekly Live View": "1VjeWOSX6nX_oQiU8QnB98suG2siKovuJhoBzxTK-zeI",
    "Reg Commercial Team": "10kH9Welrxx7KJEOrtfWFshOrglsG-P6vTv9otwxsPho",
    "Reg CNLS copy": "1cN29heWI-7trzBLMvEpznXslDmlCEmLHUcKDjT9uTqg",
    "Archive": "1F99kNADGaRxiuxkxG2Gq3A7Xvoh_bG0EM2C0xYi8Ocs",
}


def get_auth_headers():
    """Get authorization headers, auto-refreshing token if needed."""
    with open(TOKEN_PATH) as f:
        creds = json.load(f)

    # Check if expired
    expiry_str = creds.get("expiry", "")
    if expiry_str:
        try:
            expiry = datetime.datetime.fromisoformat(expiry_str.replace("Z", "+00:00"))
            if datetime.datetime.now(datetime.timezone.utc) >= expiry - datetime.timedelta(minutes=2):
                print("[auth] Token expired or expiring soon — refreshing...", flush=True)
                refresh_data = urllib.parse.urlencode({
                    "client_id": creds["client_id"],
                    "client_secret": creds["client_secret"],
                    "refresh_token": creds["refresh_token"],
                    "grant_type": "refresh_token",
                }).encode()
                req = urllib.request.Request(
                    creds["token_uri"], data=refresh_data,
                    headers={"Content-Type": "application/x-www-form-urlencoded"}
                )
                with urllib.request.urlopen(req, timeout=15, context=SSL_CTX) as resp:
                    token_resp = json.loads(resp.read())
                creds["token"] = token_resp["access_token"]
                new_expiry = datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(
                    seconds=token_resp.get("expires_in", 3600))
                creds["expiry"] = new_expiry.isoformat().replace("+00:00", "Z")
                with open(TOKEN_PATH, "w") as f:
                    json.dump(creds, f, indent=2)
                print("[auth] Token refreshed.", flush=True)
        except Exception as e:
            print(f"[auth] Token check/refresh failed: {e}", flush=True)

    return {"Authorization": f"Bearer {creds['token']}", "Accept": "application/json"}


def api_get(sheet_id, fields):
    """Call Google Sheets GET API with auth + SSL."""
    headers = get_auth_headers()
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}?fields={urllib.parse.quote(fields)}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def api_batch_update(sheet_id, requests_list):
    """Call Google Sheets POST batchUpdate with auth + SSL."""
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}:batchUpdate"
    body = json.dumps({"requests": requests_list}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def values_get(sheet_id, range_):
    """Read cell values from a range."""
    headers = get_auth_headers()
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{urllib.parse.quote(range_, safe='!')}"
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=SSL_CTX) as resp:
        return json.loads(resp.read())


def values_update(sheet_id, range_, values, input_option="USER_ENTERED"):
    """Write values to a range."""
    headers = get_auth_headers()
    headers["Content-Type"] = "application/json"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}/values/{urllib.parse.quote(range_, safe='!')}?valueInputOption={input_option}"
    body = json.dumps({"values": values}).encode()
    req = urllib.request.Request(url, data=body, headers=headers, method="PUT")
    with urllib.request.urlopen(req, timeout=60, context=SSL_CTX) as resp:
        return json.loads(resp.read())
