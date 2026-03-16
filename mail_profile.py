#!/usr/bin/env python3
"""
mail_profile.py

Connects to Apple/iCloud Mail via IMAP, samples your emails, extracts
personal information, and updates CLAUDE.md with a rich personal profile
so Claude Code can personalise every session.

Usage:
    python mail_profile.py

Credentials are read from a .env file (see .env.example).
"""

import imaplib
import email
import email.header
import email.utils
import os
import re
import json
import textwrap
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path

import anthropic
from dotenv import load_dotenv

load_dotenv()

# ── Config ─────────────────────────────────────────────────────────────────

IMAP_HOST = os.getenv("IMAP_HOST", "imap.mail.me.com")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
MAIL_USER = os.getenv("MAIL_USER", "")
MAIL_PASS = os.getenv("MAIL_PASS", "")  # Use an app-specific password

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")

# How many emails to sample per folder
SAMPLE_SENT   = int(os.getenv("SAMPLE_SENT", "150"))
SAMPLE_INBOX  = int(os.getenv("SAMPLE_INBOX", "100"))

CLAUDE_MD_PATH = Path(os.getenv("CLAUDE_MD_PATH", "./CLAUDE.md"))

# ── IMAP helpers ──────────────────────────────────────────────────────────

def connect() -> imaplib.IMAP4_SSL:
    if not MAIL_USER or not MAIL_PASS:
        raise ValueError(
            "MAIL_USER and MAIL_PASS must be set in your .env file.\n"
            "For iCloud, generate an app-specific password at "
            "https://appleid.apple.com/account/manage"
        )
    client = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    client.login(MAIL_USER, MAIL_PASS)
    return client


def list_folders(client: imaplib.IMAP4_SSL) -> list[str]:
    _, folders = client.list()
    names = []
    for f in folders:
        parts = f.decode().split('"')
        names.append(parts[-1].strip() if len(parts) >= 2 else f.decode().split()[-1])
    return names


def find_folder(client: imaplib.IMAP4_SSL, candidates: list[str]) -> str | None:
    available = list_folders(client)
    for candidate in candidates:
        for folder in available:
            if candidate.lower() in folder.lower():
                return folder
    return None


def fetch_messages(client: imaplib.IMAP4_SSL, folder: str, limit: int) -> list[email.message.Message]:
    client.select(f'"{folder}"', readonly=True)
    _, data = client.search(None, "ALL")
    ids = data[0].split()
    # Take the most recent `limit` messages
    ids = ids[-limit:]
    messages = []
    for uid in ids:
        _, msg_data = client.fetch(uid, "(RFC822)")
        for part in msg_data:
            if isinstance(part, tuple):
                msg = email.message_from_bytes(part[1])
                messages.append(msg)
    return messages


# ── Email parsing helpers ──────────────────────────────────────────────────

def decode_header_value(value: str | None) -> str:
    if not value:
        return ""
    parts = email.header.decode_header(value)
    decoded = []
    for b, charset in parts:
        if isinstance(b, bytes):
            decoded.append(b.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(b)
    return "".join(decoded)


def extract_address(raw: str) -> tuple[str, str]:
    """Return (display_name, email_address) from a raw header value."""
    name, addr = email.utils.parseaddr(decode_header_value(raw))
    return name.strip(), addr.strip().lower()


def get_text_body(msg: email.message.Message) -> str:
    """Extract plain-text body, falling back to HTML stripped of tags."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    body += payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                    break
            elif ct == "text/html" and not body:
                payload = part.get_payload(decode=True)
                if payload:
                    html = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
                    body = re.sub(r"<[^>]+>", " ", html)
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
    return body.strip()


# ── Structured data extraction ─────────────────────────────────────────────

def collect_signals(sent: list, inbox: list) -> dict:
    """
    Parse emails and return structured signals for profile building.
    """
    my_addresses: set[str] = set()
    contact_counter: Counter = Counter()
    domain_counter: Counter = Counter()
    subject_words: Counter = Counter()
    signatures: list[str] = []
    body_samples: list[str] = []

    # ── Sent mail: richest source about the user themselves ──────────────
    for msg in sent:
        from_name, from_addr = extract_address(msg.get("From", ""))
        if from_addr:
            my_addresses.add(from_addr)

        # Recipient contacts
        for header in ("To", "Cc"):
            raw = msg.get(header, "")
            for part in raw.split(","):
                name, addr = extract_address(part)
                if addr:
                    contact_counter[addr] += 1
                    domain = addr.split("@")[-1] if "@" in addr else ""
                    if domain:
                        domain_counter[domain] += 1

        # Subject keywords
        subj = decode_header_value(msg.get("Subject", ""))
        for word in re.findall(r"\b[A-Za-z]{4,}\b", subj):
            subject_words[word.lower()] += 1

        # Body sample + signature
        body = get_text_body(msg)
        if body:
            body_samples.append(body[:800])
            # Look for email signatures (heuristic: last 20 lines, starts with -- or contains phone)
            lines = body.strip().splitlines()
            sig_lines = []
            for i, line in enumerate(reversed(lines[-25:])):
                if re.search(r"--|^\s*[\w\s]+\||\+\d[\d\s\-()]{6,}|[Tt]itle|[Rr]egards|[Ss]incerely", line):
                    sig_lines = lines[-(i + 8):]
                    break
            if sig_lines:
                signatures.append("\n".join(sig_lines[:10]))

    # ── Inbox: who contacts the user ─────────────────────────────────────
    for msg in inbox:
        from_name, from_addr = extract_address(msg.get("From", ""))
        if from_addr:
            domain = from_addr.split("@")[-1] if "@" in from_addr else ""
            if domain and domain not in ("gmail.com", "yahoo.com", "hotmail.com",
                                         "outlook.com", "icloud.com", "me.com"):
                domain_counter[domain] += 1

    return {
        "my_addresses": sorted(my_addresses),
        "top_contacts": contact_counter.most_common(30),
        "top_domains": domain_counter.most_common(20),
        "subject_keywords": subject_words.most_common(40),
        "signatures": signatures[:6],
        "body_samples": body_samples[:20],
    }


# ── Claude profile synthesis ───────────────────────────────────────────────

SYSTEM_PROMPT = textwrap.dedent("""
    You are a personal assistant helping Claude Code learn about its user so
    future sessions can be more personalised and efficient.

    You will receive structured signals extracted from the user's email.
    Synthesise them into a concise personal profile in Markdown.

    Output **only** the Markdown section content (no fences, no preamble).
    Use these headings exactly:
    - ## Identity
    - ## Professional context
    - ## Frequent collaborators
    - ## Recurring topics & interests
    - ## Communication style
    - ## Preferences & working style

    Under each heading write 3–6 bullet points. Be specific where the data
    supports it. Do not invent facts not supported by the signals.
    Never include raw email content, passwords, or sensitive financial data.
""").strip()


def build_profile_with_claude(signals: dict) -> str:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Prepare a sanitised summary (avoid sending full bodies)
    summary = json.dumps({
        "email_addresses": signals["my_addresses"],
        "top_contacts": signals["top_contacts"][:20],
        "top_domains": signals["top_domains"][:15],
        "subject_keywords": signals["subject_keywords"][:30],
        "signature_samples": signals["signatures"][:4],
        "body_snippets": [s[:300] for s in signals["body_samples"][:10]],
    }, indent=2)

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2000,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": (
                    "Here are the signals extracted from my email.\n"
                    "Please build my personal profile.\n\n"
                    f"```json\n{summary}\n```"
                ),
            }
        ],
    )
    return message.content[0].text.strip()


# ── CLAUDE.md update ───────────────────────────────────────────────────────

PROFILE_START = "<!-- mail-profile:start -->"
PROFILE_END   = "<!-- mail-profile:end -->"


def update_claude_md(profile_markdown: str) -> None:
    existing = CLAUDE_MD_PATH.read_text(encoding="utf-8") if CLAUDE_MD_PATH.exists() else ""

    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    block = (
        f"{PROFILE_START}\n"
        f"# Personal profile\n"
        f"_Last updated: {timestamp}_\n\n"
        f"{profile_markdown}\n"
        f"{PROFILE_END}"
    )

    if PROFILE_START in existing and PROFILE_END in existing:
        # Replace existing block
        pattern = re.compile(
            re.escape(PROFILE_START) + r".*?" + re.escape(PROFILE_END),
            re.DOTALL,
        )
        updated = pattern.sub(block, existing)
    else:
        # Append to end
        updated = existing.rstrip() + "\n\n" + block + "\n"

    CLAUDE_MD_PATH.write_text(updated, encoding="utf-8")
    print(f"  CLAUDE.md updated at {CLAUDE_MD_PATH.resolve()}")


# ── Main ───────────────────────────────────────────────────────────────────

def main() -> None:
    print("Connecting to mail server …")
    client = connect()

    # Locate Sent folder (iCloud names it "Sent Messages")
    sent_folder = find_folder(client, ["Sent Messages", "Sent", "INBOX.Sent"])
    if not sent_folder:
        print("  Warning: could not find Sent folder. Falling back to INBOX only.")
        sent_msgs = []
    else:
        print(f"  Fetching up to {SAMPLE_SENT} sent messages from '{sent_folder}' …")
        sent_msgs = fetch_messages(client, sent_folder, SAMPLE_SENT)
        print(f"  Retrieved {len(sent_msgs)} sent messages.")

    inbox_folder = find_folder(client, ["INBOX", "Inbox"])
    if not inbox_folder:
        print("  Warning: could not find INBOX.")
        inbox_msgs = []
    else:
        print(f"  Fetching up to {SAMPLE_INBOX} inbox messages from '{inbox_folder}' …")
        inbox_msgs = fetch_messages(client, inbox_folder, SAMPLE_INBOX)
        print(f"  Retrieved {len(inbox_msgs)} inbox messages.")

    client.logout()

    print("Extracting signals …")
    signals = collect_signals(sent_msgs, inbox_msgs)
    print(f"  Found {len(signals['my_addresses'])} own address(es), "
          f"{len(signals['top_contacts'])} contacts, "
          f"{len(signals['signatures'])} signature sample(s).")

    print("Building profile with Claude …")
    profile_md = build_profile_with_claude(signals)

    print("Updating CLAUDE.md …")
    update_claude_md(profile_md)

    print("\nDone. Your personal profile has been written to CLAUDE.md.")
    print("Claude Code will read it automatically at the start of every session.")


if __name__ == "__main__":
    main()
