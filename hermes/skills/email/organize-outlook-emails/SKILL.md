---
name: organize-outlook-emails
description: Organize unread Outlook emails by topic/subject and move to folders using himalaya CLI.
version: 1.0.0
author: community
license: MIT
metadata:
  hermes:
    tags: [Email, Outlook, IMAP, Organization]
prerequisites:
  commands: [himalaya, python3]
---

# Organize Outlook Emails by Topic

This skill organizes unread Outlook emails by analyzing their subjects and moving them to topic‑based folders.

## Prerequisites

1. **Himalaya CLI** installed and configured for your Outlook account.
2. **Python 3** installed (for the categorization script).
3. **IMAP must be enabled** in your Outlook.com settings:
   - Go to [Outlook.com](https://outlook.com) > Settings (gear) > View all Outlook settings > Mail > Sync email
   - Turn **ON** "Let devices and apps use IMAP"
   - Save

## Configuration

### Step 1: IMAP vs OAuth2 — Know Which You Need

Microsoft has been **phasing out basic password authentication** for IMAP. Before configuring:

- **Personal Outlook.com accounts** — basic auth with an app password *may* still work if IMAP is enabled (see prerequisite above).
- **Office 365 work/school accounts** — basic auth is **blocked**; you *must* use OAuth2.

**Quick test** — if `himalaya folder list` returns `"AUTHENTICATE failed."`, then basic auth is blocked and you need OAuth2.

### Step 2a: Basic Auth Setup (if IMAP auth works)

Create `~/.config/himalaya/config.toml`:

```toml
[accounts.outlook]
email = "your-email@outlook.com"
display-name = "Your Name"
default = true

backend.type = "imap"
backend.host = "outlook.office365.com"
backend.port = 993
backend.encryption.type = "tls"
backend.login = "your-email@outlook.com"
backend.auth.type = "password"
backend.auth.cmd = "security find-generic-password -a your-email@outlook.com -s himalaya-imap -w"

message.send.backend.type = "smtp"
message.send.backend.host = "smtp.office365.com"
message.send.backend.port = 587
message.send.backend.encryption.type = "start-tls"
message.send.backend.login = "your-email@outlook.com"
message.send.backend.auth.type = "password"
message.send.backend.auth.cmd = "security find-generic-password -a your-email@outlook.com -s himalaya-imap -w"

[accounts.outlook.folder.alias]
inbox = "INBOX"
sent = "Sent Items"
drafts = "Drafts"
trash = "Deleted Items"
```

Store password in macOS keychain:

```bash
security add-generic-password -a "your-email@outlook.com" -s "himalaya-imap" -w "YOUR_APP_PASSWORD"
```

### Step 2b: OAuth2 Setup (when basic auth is blocked)

If you get `"AUTHENTICATE failed"`, you need OAuth2. This requires an Azure AD app registration:

1. Go to https://portal.azure.com > Azure Active Directory > App registrations > New registration
   - Name: `himalaya`
   - Account type: "Accounts in any organizational directory and personal Microsoft accounts"
   - Redirect URI: `http://localhost:7892` (type: Web)
   - Register, then note the **Application (client) ID**

2. Under "Certificates & secrets" > "Client secrets" > New client secret, note the secret value.

3. Under "API permissions" > Add a permission > Microsoft Graph > Delegated permissions:
   - Add `IMAP.AccessAsUser.All`, `Mail.ReadWrite`, `Mail.Send`, `offline_access`
   - Grant admin consent (or user consent at runtime)

4. Configure himalaya with OAuth2:

```toml
[accounts.outlook]
email = "your-email@outlook.com"
display-name = "Your Name"
default = true

backend.type = "imap"
backend.host = "outlook.office365.com"
backend.port = 993
backend.encryption.type = "tls"
backend.login = "your-email@outlook.com"
backend.auth.type = "oauth2"
backend.auth.client-id = "YOUR_CLIENT_ID"
backend.auth.client-secret.cmd = "security find-generic-password -a oauth -s himalaya-client-secret -w"
backend.auth.tenant = "common"
backend.auth.auth-url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
backend.auth.token-url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"

message.send.backend.type = "smtp"
message.send.backend.host = "smtp.office365.com"
message.send.backend.port = 587
message.send.backend.encryption.type = "start-tls"
message.send.backend.login = "your-email@outlook.com"
message.send.backend.auth.type = "oauth2"
message.send.backend.auth.client-id = "YOUR_CLIENT_ID"
message.send.backend.auth.client-secret.cmd = "security find-generic-password -a oauth -s himalaya-client-secret -w"
message.send.backend.auth.tenant = "common"
message.send.backend.auth.auth-url = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
message.send.backend.auth.token-url = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
```

Store the client secret in keychain:

```bash
security add-generic-password -a "oauth" -s "himalaya-client-secret" -w "YOUR_CLIENT_SECRET"
```

### Step 3: Test Connectivity

```bash
himalaya folder list
```

Expected output includes `INBOX`, `Sent Items`, `Drafts`, `Deleted Items`.

If this fails, see Troubleshooting below.

## How It Works

The script:
1. Fetches unread emails from the INBOX (or any folder you specify).
2. Extracts subject, sender, and message ID.
3. Matches the subject against a keyword‑to‑folder mapping.
4. Creates missing folders (if needed).
5. Moves each email to its assigned folder.

## Keyword Mapping

Edit the `KEYWORD_MAP` dictionary in the script to match your own topics. Example:

```python
KEYWORD_MAP = {
    "work": ["meeting", "report", "deadline"],
    "personal": ["family", "friends", "vacation"],
    "finance": ["invoice", "bill", "payment"],
    "newsletter": ["newsletter", "subscription", "digest"],
    "shopping": ["order", "receipt", "amazon"],
}
```

If no keywords match, the email stays in INBOX (or you can define a default folder).

## Usage

### 1. Save the Script

Save the script below as `organize_emails.py` in a convenient location.

### 2. Run the Script

```bash
python3 organize_emails.py
```

The script will show a preview of moves and ask for confirmation before making changes.

### 3. Automate (Optional)

Add a cron job to run the script periodically:

```bash
crontab -e
# Add line: 0 */2 * * * /usr/bin/python3 /path/to/organize_emails.py
```

## Script

```python
#!/usr/bin/env python3
"""
Organize unread Outlook emails by subject keywords.
Requires himalaya CLI configured and accessible.
"""

import subprocess
import json
import sys
import re
from typing import List, Dict, Optional

# --- Configuration ---
# Map folder names to lists of keyword patterns (case‑insensitive)
KEYWORD_MAP = {
    "Work": ["meeting", "report", "deadline", "project", "team"],
    "Personal": ["family", "friends", "vacation", "party", "dinner"],
    "Finance": ["invoice", "bill", "payment", "bank", "statement"],
    "Newsletters": ["newsletter", "digest", "subscription", "weekly"],
    "Shopping": ["order", "receipt", "amazon", "delivery", "tracking"],
    "Travel": ["flight", "hotel", "booking", "itinerary"],
}

# If no keyword matches, leave in INBOX (set to None) or move to a default folder
DEFAULT_FOLDER = None  # or "Uncategorized"

# Himalaya account name (as defined in config.toml). Use None for default.
ACCOUNT = None

# Number of unread emails to process (0 = all)
LIMIT = 50

# --- Himalaya helpers ---

def himalaya_cmd(args: List[str]) -> str:
    """Run himalaya command and return stdout."""
    cmd = ["himalaya"]
    if ACCOUNT:
        cmd.extend(["--account", ACCOUNT])
    cmd.extend(args)
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running himalaya: {result.stderr}")
        sys.exit(1)
    return result.stdout

def get_unread_envelopes(limit: int = 0) -> List[Dict]:
    """Fetch unread emails as JSON envelopes."""
    args = ["envelope", "list", "--output", "json"]
    if limit > 0:
        args.extend(["--page-size", str(limit)])
    # Filter by flag "unread" – himalaya doesn't have a built‑in unread filter,
    # so we fetch all and filter later.
    output = himalaya_cmd(args)
    try:
        envelopes = json.loads(output)
    except json.JSONDecodeError:
        print("Failed to parse himalaya output as JSON.")
        sys.exit(1)
    # Filter for unread (flags does not contain "Seen")
    unread = [e for e in envelopes if "Seen" not in e.get("flags", [])]
    return unread

def get_folders() -> List[str]:
    """List existing folder names."""
    output = himalaya_cmd(["folder", "list"])
    # Output is one folder per line, optionally with hierarchy.
    folders = [line.strip() for line in output.splitlines() if line.strip()]
    return folders

def create_folder(folder: str) -> bool:
    """Create a folder if it doesn't exist."""
    folders = get_folders()
    if folder in folders:
        return True
    print(f"Creating folder: {folder}")
    result = subprocess.run(
        ["himalaya", "folder", "create", folder],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        print(f"Failed to create folder {folder}: {result.stderr}")
        return False
    return True

def move_message(msg_id: str, folder: str) -> bool:
    """Move a message to a folder."""
    result = subprocess.run(
        ["himalaya", "message", "move", msg_id, folder],
        capture_output=True,
        text=True
    )
    if result.returncode != 0:
        print(f"Failed to move message {msg_id} to {folder}: {result.stderr}")
        return False
    return True

# --- Categorization ---

def match_folder(subject: str) -> Optional[str]:
    """Return folder name based on subject keywords."""
    subject_lower = subject.lower()
    for folder, keywords in KEYWORD_MAP.items():
        for kw in keywords:
            if re.search(rf'\b{re.escape(kw.lower())}\b', subject_lower):
                return folder
    return DEFAULT_FOLDER

def main():
    print("Fetching unread emails...")
    envelopes = get_unread_envelopes(LIMIT)
    if not envelopes:
        print("No unread emails found.")
        return

    print(f"Found {len(envelopes)} unread email(s).")
    # Get existing folders
    existing_folders = get_folders()
    print(f"Existing folders: {', '.join(existing_folders)}")

    # Plan moves
    moves = []
    for env in envelopes:
        msg_id = str(env["id"])
        subject = env.get("subject", "(no subject)")
        from_addr = env.get("from", {}).get("name", env.get("from", {}).get("addr", "unknown"))
        folder = match_folder(subject)
        if folder:
            moves.append((msg_id, subject[:50], from_addr, folder))

    if not moves:
        print("No emails matched any keyword. Nothing to do.")
        return

    print("\nPlanned moves:")
    for i, (msg_id, subj, from_addr, folder) in enumerate(moves, 1):
        print(f"{i:2}. {subj} ({from_addr}) → {folder}")

    # Confirm
    response = input("\nProceed with moving? (y/N): ").strip().lower()
    if response != "y":
        print("Aborted.")
        return

    # Ensure folders exist
    folders_needed = {folder for _, _, _, folder in moves}
    for folder in folders_needed:
        if folder not in existing_folders:
            if not create_folder(folder):
                print(f"Aborting because folder {folder} could not be created.")
                return

    # Perform moves
    success = 0
    for msg_id, subj, _, folder in moves:
        if move_message(msg_id, folder):
            success += 1
            print(f"Moved: {subj} → {folder}")
        else:
            print(f"Failed: {subj}")

    print(f"\nDone. Successfully moved {success} of {len(moves)} emails.")

if __name__ == "__main__":
    main()
```

## Troubleshooting

- **`"AUTHENTICATE failed."`** — Microsoft has blocked basic password auth for this account. Either:
  - Enable IMAP in Outlook settings (Settings > Mail > Sync email > "Let devices and apps use IMAP") and retry, OR
  - Switch to **OAuth2** setup (see Step 2b above) — required for work/school accounts and some personal accounts.
- **`"cannot get imap password"`**: Ensure your password is stored in the keychain and the `auth.cmd` points to the correct service/account name.
- **`"The specified item already exists in the keychain"`**: Run `security delete-generic-password -a "your-email@outlook.com" -s "himalaya-imap"` first, then add again.
- **`"folder not found"`**: Check folder aliases in config.toml. Outlook uses "Sent Items", "Deleted Items", etc.
- **No emails moved**: Adjust keyword patterns or add more keywords.
- **Himalaya command not found**: Install himalaya (`brew install himalaya`).

## Customization

- Modify `KEYWORD_MAP` to fit your own topics.
- Change `DEFAULT_FOLDER` to automatically categorize unmatched emails.
- Adjust `LIMIT` to process only recent emails.

## Notes

- The script only moves emails; it does not delete or mark as read.
- Run the script manually first to verify the planned moves.
- For large mailboxes, consider increasing `LIMIT` gradually.

## Future Enhancements

- Use LLM to categorize emails based on full content.
- Support multiple email accounts.
- Add logging and error recovery.
- Integrate with Hermes Agent as a skill with interactive configuration.