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