#!/usr/bin/env python3
"""
Mail Analysis for Claude Personalization
=========================================
Analyzes your email metadata to build a personal profile and writes it to
CLAUDE.md so Claude can provide more personalized assistance.

Supports two mail backends:
  - Gmail API (OAuth2, recommended — reads metadata only, no message bodies by default)
  - IMAP (works with any provider using app passwords)

Usage:
  python analyze_mail.py --source gmail          # Gmail API (needs credentials.json)
  python analyze_mail.py --source imap \\
      --host imap.gmail.com --user you@gmail.com  # IMAP

Gmail API Setup:
  1. Go to https://console.cloud.google.com/
  2. Create a project, enable the Gmail API
  3. Create OAuth 2.0 credentials (Desktop App type)
  4. Download the JSON file and save as credentials.json in this directory
  5. Run: python analyze_mail.py --source gmail
     (browser opens for one-time authorisation; token is cached for future runs)

Privacy note:
  By default only email headers (From, To, Subject, Date) and short snippets
  are read.  No raw message bodies leave your machine.  The summary sent to
  Claude contains aggregated patterns, not individual messages.
"""

import os
import re
import json
import pickle
import imaplib
import argparse
import email as email_lib
from email.header import decode_header
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional

import anthropic

# ── Optional Gmail API dependencies ──────────────────────────────────────────
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    GMAIL_AVAILABLE = True
except ImportError:
    GMAIL_AVAILABLE = False

GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
TOKEN_FILE = Path.home() / ".claude_mail_token.pickle"
CREDENTIALS_FILE = Path("credentials.json")


# ── Gmail backend ─────────────────────────────────────────────────────────────

def setup_gmail():
    """Authenticate with the Gmail API and return a service object."""
    if not GMAIL_AVAILABLE:
        print(
            "Gmail API libraries are not installed.\n"
            "Run:  pip install google-auth-oauthlib google-api-python-client"
        )
        return None

    creds: Optional[Credentials] = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as fh:
            creds = pickle.load(fh)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                print(
                    "\nGmail API credentials not found.\n\n"
                    "Setup steps:\n"
                    "  1. Visit https://console.cloud.google.com/\n"
                    "  2. Create a project and enable the Gmail API\n"
                    "  3. Create OAuth 2.0 credentials (Desktop App)\n"
                    "  4. Download the JSON and save as credentials.json here\n"
                    "  5. Re-run this script\n"
                )
                return None
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), GMAIL_SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as fh:
            pickle.dump(creds, fh)

    return build("gmail", "v1", credentials=creds)


def fetch_gmail_emails(service, days: int = 90, max_emails: int = 500) -> list[dict]:
    """Fetch email metadata from Gmail (headers + snippet only)."""
    print(f"Fetching Gmail headers for the last {days} days (max {max_emails})…")
    after_date = (datetime.now() - timedelta(days=days)).strftime("%Y/%m/%d")

    emails: list[dict] = []
    page_token = None

    while len(emails) < max_emails:
        params: dict = {
            "userId": "me",
            "q": f"after:{after_date}",
            "maxResults": min(100, max_emails - len(emails)),
        }
        if page_token:
            params["pageToken"] = page_token

        results = service.users().messages().list(**params).execute()
        messages = results.get("messages", [])
        if not messages:
            break

        for msg in messages:
            try:
                full = service.users().messages().get(
                    userId="me",
                    id=msg["id"],
                    format="metadata",
                    metadataHeaders=["From", "To", "Subject", "Date"],
                ).execute()
                hdrs = {
                    h["name"]: h["value"]
                    for h in full.get("payload", {}).get("headers", [])
                }
                emails.append(
                    {
                        "from": hdrs.get("From", ""),
                        "to": hdrs.get("To", ""),
                        "subject": hdrs.get("Subject", ""),
                        "date": hdrs.get("Date", ""),
                        "snippet": full.get("snippet", "")[:200],
                    }
                )
            except Exception:
                continue

        if len(emails) % 100 == 0 and emails:
            print(f"  … {len(emails)} fetched")

        page_token = results.get("nextPageToken")
        if not page_token:
            break

    print(f"Fetched {len(emails)} emails total.")
    return emails


# ── IMAP backend ──────────────────────────────────────────────────────────────

def _decode_header_value(raw: str | None) -> str:
    """Safely decode an RFC 2047-encoded header value."""
    if not raw:
        return ""
    parts = []
    for chunk, enc in decode_header(raw):
        if isinstance(chunk, bytes):
            parts.append(chunk.decode(enc or "utf-8", errors="replace"))
        else:
            parts.append(str(chunk))
    return " ".join(parts)


def fetch_imap_emails(
    host: str,
    username: str,
    password: str,
    days: int = 90,
    max_emails: int = 500,
    use_ssl: bool = True,
) -> list[dict]:
    """Fetch email headers via IMAP."""
    print(f"Connecting to {host} as {username}…")
    conn = imaplib.IMAP4_SSL(host) if use_ssl else imaplib.IMAP4(host)
    conn.login(username, password)
    conn.select("INBOX")

    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    _, data = conn.search(None, f"SINCE {since}")
    ids = data[0].split()
    # Take the most-recent max_emails
    ids = ids[-max_emails:] if len(ids) > max_emails else ids

    emails: list[dict] = []
    for i, msg_id in enumerate(reversed(ids)):
        try:
            _, raw_data = conn.fetch(
                msg_id, "(BODY.PEEK[HEADER.FIELDS (FROM TO SUBJECT DATE)])"
            )
            msg = email_lib.message_from_bytes(raw_data[0][1])
            emails.append(
                {
                    "from": _decode_header_value(msg.get("From")),
                    "to": _decode_header_value(msg.get("To")),
                    "subject": _decode_header_value(msg.get("Subject")),
                    "date": msg.get("Date", ""),
                    "snippet": "",
                }
            )
        except Exception:
            continue

        if (i + 1) % 100 == 0:
            print(f"  … {i + 1} fetched")

    conn.logout()
    print(f"Fetched {len(emails)} emails total.")
    return emails


# ── Summary builder ───────────────────────────────────────────────────────────

def _normalize_subject(subject: str) -> str:
    """Strip common prefixes and normalize for deduplication."""
    s = re.sub(r"^(Re|Fwd?|AW|WG|SV|TR):\s*", "", subject, flags=re.IGNORECASE).strip()
    return s.lower()


def prepare_email_summary(emails: list[dict]) -> str:
    """
    Aggregate email metadata into a compact, privacy-conscious summary
    for Claude to analyse.  Individual messages are never included verbatim.
    """
    sender_counts: dict[str, int] = {}
    domain_counts: dict[str, int] = {}
    subjects: list[str] = []
    snippets: list[tuple[str, str]] = []

    for e in emails:
        from_raw = e.get("from", "")
        if "@" in from_raw:
            domain = from_raw.split("@")[-1].split(">")[0].strip().lower()
            # Strip trailing > from "Name <user@domain>" patterns
            domain = domain.rstrip(">").strip()
            domain_counts[domain] = domain_counts.get(domain, 0) + 1
        if from_raw:
            sender_counts[from_raw] = sender_counts.get(from_raw, 0) + 1

        subj = e.get("subject", "").strip()
        if subj:
            subjects.append(subj)

        snippet = e.get("snippet", "").strip()
        if snippet:
            snippets.append((subj, snippet))

    top_senders = sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)[:40]
    top_domains = sorted(domain_counts.items(), key=lambda x: x[1], reverse=True)[:25]

    # Unique subjects (deduped by normalized form)
    seen: set[str] = set()
    unique_subjects: list[str] = []
    for s in subjects:
        norm = _normalize_subject(s)
        if norm and norm not in seen:
            seen.add(norm)
            unique_subjects.append(s)
        if len(unique_subjects) >= 200:
            break

    lines: list[str] = [
        f"# Email Pattern Analysis  ({len(emails)} emails, last analysis window)",
        "",
        "## Top senders (name + address, count)",
    ]
    for addr, count in top_senders:
        lines.append(f"  {count:4d}×  {addr}")

    lines += ["", "## Top sending domains (count)"]
    for domain, count in top_domains:
        lines.append(f"  {count:4d}×  {domain}")

    lines += ["", "## Sample unique email subjects"]
    for s in unique_subjects[:150]:
        lines.append(f"  • {s}")

    if snippets:
        lines += ["", "## Sample message snippets (first 200 chars each)"]
        for subj, snip in snippets[:40]:
            lines.append(f"  [{subj}]  {snip}")

    return "\n".join(lines)


# ── Claude analysis ───────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are helping to build a personalization profile for Claude by analysing
anonymised email metadata (senders, subjects, snippets — no full bodies).

Your job is to infer who this person is, what they work on, and how Claude
can serve them better.  Be specific and concrete wherever the data supports
it, but do not over-infer or fabricate details.

Do NOT include raw email addresses, phone numbers, or other PII.
Write in the second person ("You work in…", "Your team…") so the profile
reads naturally as a Claude memory file.
"""

USER_PROMPT_TEMPLATE = """\
Please analyse the email patterns below and produce a `## Personal Context`
section for a CLAUDE.md personalisation file.

{existing_note}

The section should cover:

1. **Who you are** – role, industry, organisation type (inferred from patterns)
2. **Work context** – recurring projects, topics, and technology domains
3. **Key relationships** – types of people and organisations you interact with
4. **Interests & expertise** – topics that surface repeatedly
5. **Communication style** – how you tend to communicate (if inferable)
6. **Suggestions for Claude** – concrete ways Claude should adapt its responses
   (e.g. assume familiarity with X, prefer concise answers, use British English…)
7. **Signals** – timezone, location, or schedule clues if present

Format as clean Markdown.  Start *exactly* with `## Personal Context`.
Be actionable and specific; avoid vague generalisations.

---

{summary}
"""


def analyze_with_claude(email_summary: str, existing_claude_md: str = "") -> str:
    """Stream Claude's analysis of the email summary and return the result."""
    client = anthropic.Anthropic()

    existing_note = (
        f"Existing CLAUDE.md content is provided for context; "
        f"merge and update rather than duplicate:\n\n{existing_claude_md}"
        if existing_claude_md.strip()
        else ""
    )

    prompt = USER_PROMPT_TEMPLATE.format(
        existing_note=existing_note, summary=email_summary
    )

    print("\nAnalysing with Claude (streaming)…\n")
    print("─" * 60)

    full_response = ""
    with client.messages.stream(
        model="claude-opus-4-6",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        for text in stream.text_stream:
            print(text, end="", flush=True)
            full_response += text

    print("\n" + "─" * 60 + "\n")
    return full_response


# ── CLAUDE.md writer ──────────────────────────────────────────────────────────

SECTION_RE = re.compile(
    r"<!--\s*Auto-generated.*?-->\s*## Personal Context.*?(?=\n## |\Z)",
    re.DOTALL,
)


def update_claude_md(profile: str, claude_md_path: Path) -> None:
    """Insert or replace the Personal Context section in CLAUDE.md."""
    existing = claude_md_path.read_text() if claude_md_path.exists() else ""

    # Strip any previous auto-generated section
    cleaned = SECTION_RE.sub("", existing).rstrip()

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    injected = (
        f"\n\n<!-- Auto-generated from mail analysis on {timestamp} -->\n{profile}"
    )

    new_content = (cleaned + injected).strip() + "\n"
    claude_md_path.write_text(new_content)
    print(f"✓  Written to {claude_md_path}")


# ── CLI ───────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Analyse your mail to personalise Claude",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--source",
        choices=["gmail", "imap"],
        default="gmail",
        help="Mail backend to use (default: gmail)",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=90,
        help="Days of history to analyse (default: 90)",
    )
    parser.add_argument(
        "--max-emails",
        type=int,
        default=500,
        help="Maximum number of emails to fetch (default: 500)",
    )
    # IMAP options
    parser.add_argument(
        "--host",
        default="imap.gmail.com",
        help="IMAP hostname (default: imap.gmail.com)",
    )
    parser.add_argument("--user", help="IMAP username / email address")
    parser.add_argument(
        "--password",
        help="IMAP password (or set MAIL_PASSWORD env var)",
    )
    parser.add_argument(
        "--no-ssl",
        action="store_true",
        help="Disable TLS/SSL for IMAP (not recommended)",
    )
    # Output
    parser.add_argument(
        "--output",
        default="CLAUDE.md",
        help="Path to CLAUDE.md (default: CLAUDE.md in current directory)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print the generated profile without writing CLAUDE.md",
    )

    args = parser.parse_args()

    # ── Fetch emails ──
    if args.source == "gmail":
        service = setup_gmail()
        if service is None:
            return 1
        emails = fetch_gmail_emails(
            service, days=args.days, max_emails=args.max_emails
        )
    else:
        user = args.user or os.environ.get("MAIL_USER", "")
        password = args.password or os.environ.get("MAIL_PASSWORD", "")
        if not user or not password:
            print(
                "IMAP requires a username and password.\n"
                "Pass --user / --password or set MAIL_USER / MAIL_PASSWORD."
            )
            return 1
        emails = fetch_imap_emails(
            args.host,
            user,
            password,
            days=args.days,
            max_emails=args.max_emails,
            use_ssl=not args.no_ssl,
        )

    if not emails:
        print("No emails found in the specified window.")
        return 1

    # ── Build summary ──
    print("Preparing anonymised summary…")
    summary = prepare_email_summary(emails)

    # ── Read existing CLAUDE.md for context ──
    claude_md_path = Path(args.output)
    existing_md = claude_md_path.read_text() if claude_md_path.exists() else ""

    # ── Analyse with Claude ──
    profile = analyze_with_claude(summary, existing_md)

    if args.dry_run:
        print("\n[Dry run — CLAUDE.md not modified]\n")
        print(profile)
        return 0

    # ── Write CLAUDE.md ──
    update_claude_md(profile, claude_md_path)
    print(
        "\nPersonalisation profile saved.\n"
        "Claude will use this context in future conversations in this directory.\n"
        "\nTip: commit CLAUDE.md to your repo so the profile persists across machines."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
