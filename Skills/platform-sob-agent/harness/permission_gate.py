"""
permission_gate.py — 操作许可
Sends a WeChat notification summarising the agent's intended action,
then waits for the user's explicit approval before proceeding.

ALL steps require approval — reads as well as writes.

WeChat delivery is handled via a pluggable notifier adapter.
Configure your preferred channel in wechat_config.json (see below).
"""

import json
import sys
import time
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
AGENT_DIR = Path(__file__).parent.parent
CONFIG_FILE = AGENT_DIR / "wechat_config.json"
SHANGHAI_TZ = timezone(timedelta(hours=8))

# Default timeout: wait up to 30 minutes for user reply
APPROVAL_TIMEOUT_SECONDS = 30 * 60

"""
wechat_config.json format (create this file manually, do NOT commit to git):

Option A — WxPusher (personal WeChat via subscription account):
{
  "method": "wxpusher",
  "app_token": "AT_XXXXXXXXXXXXXXXX",
  "uid": "UID_XXXXXXXXXXXXXXXX"
}

Option B — WeCom (企业微信) webhook:
{
  "method": "wecom_webhook",
  "webhook_url": "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=XXXX"
}

Option C — Terminal only (for testing, no WeChat):
{
  "method": "terminal"
}
"""


# ── Notifier adapters ────────────────────────────────────────────────────────
def _load_config() -> dict:
    if not CONFIG_FILE.exists():
        print(f"[permission_gate] WARNING: {CONFIG_FILE} not found. Falling back to terminal mode.")
        return {"method": "terminal"}
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def _send_wxpusher(message: str, config: dict):
    """Send via WxPusher — personal WeChat subscription account."""
    import urllib.request
    payload = json.dumps({
        "appToken": config["app_token"],
        "content": message,
        "contentType": 1,        # 1 = plain text
        "uids": [config["uid"]],
    }).encode("utf-8")
    req = urllib.request.Request(
        "https://wxpusher.zjiecode.com/api/send/message",
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        result = json.loads(resp.read())
        if result.get("success"):
            print("[permission_gate] WeChat message sent via WxPusher.")
        else:
            print(f"[permission_gate] WxPusher error: {result}")


def _send_wecom_webhook(message: str, config: dict):
    """Send via WeCom (企业微信) webhook."""
    import urllib.request
    payload = json.dumps({
        "msgtype": "text",
        "text": {"content": message}
    }).encode("utf-8")
    req = urllib.request.Request(
        config["webhook_url"],
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=10) as resp:
        result = json.loads(resp.read())
        print(f"[permission_gate] WeCom webhook response: {result}")


def _send_terminal(message: str, config: dict):
    """Print to terminal — for local testing without WeChat."""
    print("\n" + "═" * 60)
    print("📋 AGENT ACTION PLAN — APPROVAL REQUIRED")
    print("═" * 60)
    print(message)
    print("═" * 60)


def send_wechat(message: str):
    """Send a WeChat notification using the configured method."""
    config = _load_config()
    method = config.get("method", "terminal")
    try:
        if method == "wxpusher":
            _send_wxpusher(message, config)
        elif method == "wecom_webhook":
            _send_wecom_webhook(message, config)
        else:
            _send_terminal(message, config)
    except Exception as e:
        print(f"[permission_gate] Failed to send notification: {e}")
        print("[permission_gate] Falling back to terminal display.")
        _send_terminal(message, config)


# ── Approval gate ────────────────────────────────────────────────────────────
def request_approval(
    step_name: str,
    action_summary: str,
    sandbox_mode: bool = False
) -> bool:
    """
    Send an approval request to the user and wait for their reply.

    In sandbox mode: auto-approves and logs [DRY_RUN] instead of waiting.

    Returns True if approved, False if rejected or timed out.
    """
    ts = datetime.now(SHANGHAI_TZ).strftime("%Y-%m-%d %H:%M")

    if sandbox_mode:
        print(f"[permission_gate] [DRY_RUN] Step {step_name} — auto-approved in sandbox mode.")
        print(f"[permission_gate] [DRY_RUN] Action summary:\n{action_summary}")
        return True

    message = (
        f"🤖 Platform SOB Agent — {ts}\n\n"
        f"📌 Step: {step_name}\n\n"
        f"{action_summary}\n\n"
        f"Reply YES to approve, NO to halt."
    )
    send_wechat(message)

    # Wait for terminal input (for WeChat-integrated setups, this would
    # instead poll a reply endpoint or check a shared flag file)
    print(f"\n[permission_gate] Waiting for approval for Step {step_name}...")
    print("[permission_gate] Type 'yes' to approve or 'no' to halt:")

    start = time.time()
    while time.time() - start < APPROVAL_TIMEOUT_SECONDS:
        try:
            reply = input(">>> ").strip().lower()
        except EOFError:
            reply = ""

        if reply in ("yes", "y", "是", "好", "ok", "approve"):
            print(f"[permission_gate] ✅ Step {step_name} approved.")
            return True
        elif reply in ("no", "n", "否", "halt", "stop", "reject"):
            print(f"[permission_gate] ❌ Step {step_name} rejected by user.")
            return False
        else:
            print("[permission_gate] Please type 'yes' or 'no'.")

    # Timed out
    send_wechat(
        f"⚠️ Platform SOB Agent — Approval timeout\n"
        f"Step {step_name} received no reply within {APPROVAL_TIMEOUT_SECONDS // 60} minutes.\n"
        f"Agent has halted."
    )
    print(f"[permission_gate] ⏱️ Approval timeout for Step {step_name}. Halting.")
    return False


def alert_failure(step_name: str, detail: str):
    """Send an alert when a validation step fails (e.g. Step 4 non-zero values)."""
    ts = datetime.now(SHANGHAI_TZ).strftime("%Y-%m-%d %H:%M")
    message = (
        f"⚠️ Platform SOB Agent — {ts}\n\n"
        f"❌ Step {step_name} FAILED\n\n"
        f"{detail}\n\n"
        f"Archive step has been skipped. Please check the sheet and advise."
    )
    send_wechat(message)
    print(f"[permission_gate] Alert sent for failed step {step_name}.")


def send_outcome_report(status: str, summary: str, audit_log_path: str):
    """Send the final outcome report via WeChat after every run."""
    ts = datetime.now(SHANGHAI_TZ).strftime("%Y-%m-%d %H:%M")
    icon = "✅" if "Success" in status else ("⚠️" if "Partial" in status else "❌")
    message = (
        f"{icon} Platform SOB Agent — Run Complete ({ts})\n\n"
        f"Status: {status}\n\n"
        f"{summary}\n\n"
        f"Audit log: {audit_log_path}"
    )
    send_wechat(message)


# ── CLI test ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    approved = request_approval(
        step_name="Step 2 — Duplicate ADG ADO Tab",
        action_summary=(
            "I will duplicate the most recent tab '26/05/07 ADG ADO' in [Reg CNLS copy]\n"
            "and name it '26/05/14 ADG ADO'.\n"
            "Then paste SOB values from [Weekly Live View] red section (B5:Z62) as values only."
        )
    )
    print(f"Decision: {'Approved' if approved else 'Rejected'}")
