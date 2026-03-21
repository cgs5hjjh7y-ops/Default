#!/usr/bin/env python3
"""
RBC Investor Services — Monthly Executive Report Generator
============================================================
Generates an AI-authored executive report for the RBC Investor Services
Senior Management Team (SMT) covering the preceding calendar month.

The report is produced as a Word (.docx) document and emailed via Gmail
OAuth2 to the configured recipients.

Usage:
  python generate_report.py               # Generate & email for the previous month
  python generate_report.py --dry-run     # Generate & save locally, do not email
  python generate_report.py --year 2026 --month 2   # Target a specific month

Scheduling:
  Run via cron on the 1st of each month (see setup_cron.sh).

Gmail OAuth2 setup:
  1. Visit https://console.cloud.google.com/
  2. Create/select a project → Enable the Gmail API
  3. Create OAuth 2.0 credentials (Desktop App type)
  4. Download JSON → save as credentials.json in this directory
  5. Run once interactively so a browser token is generated:
       python generate_report.py --dry-run

Word template:
  If report_template.docx exists in this directory it will be used as the
  base document (styles, header/footer, logo, etc.).  Otherwise a clean
  document is built programmatically from the RBC brand colours.
"""

import argparse
import base64
import io
import json
import os
import pickle
import re
import sys
from calendar import month_name
from datetime import date, datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional

import anthropic
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Inches, Pt, RGBColor

# ── Configuration ─────────────────────────────────────────────────────────────

RECIPIENTS = [
    "Ian.sinclair@rbc.com",
    "ian@sinclairandsinclair.co.uk",
]

TEMPLATE_FILE = Path("report_template.docx")
CREDENTIALS_FILE = Path("credentials.json")
TOKEN_FILE = Path.home() / ".rbc_report_gmail_token.pickle"

# RBC brand colours
RBC_BLUE = RGBColor(0x00, 0x42, 0x9B)   # #00429B
RBC_GOLD = RGBColor(0xFF, 0xBE, 0x00)   # #FFBE00
RBC_DARK = RGBColor(0x1A, 0x1A, 0x2E)   # near-black

GMAIL_SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
]

# ── Lazy Gmail imports ─────────────────────────────────────────────────────────

try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build as _gs_build
    GMAIL_AVAILABLE = True
except Exception:
    GMAIL_AVAILABLE = False


# ── Date helpers ───────────────────────────────────────────────────────────────

def prev_month_info(ref: date) -> tuple[int, int]:
    """Return (year, month) for the month before ref."""
    if ref.month == 1:
        return ref.year - 1, 12
    return ref.year, ref.month - 1


def month_label(year: int, month: int) -> str:
    """E.g. 'February 2026'."""
    return f"{month_name[month]} {year}"


def next_month_label(year: int, month: int) -> str:
    if month == 12:
        return month_label(year + 1, 1)
    return month_label(year, month + 1)


# ── Claude content generation ──────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are a senior analyst at RBC Investor Services with deep expertise in
custody, fund administration, securities services, capital markets
infrastructure, and investor services regulation globally.

You produce crisp, authoritative executive briefings for the Senior Management
Team (SMT).  Your tone is professional, measured, and data-aware.  You cite
well-known industry sources (FT, Reuters, regulators, IOSCO, etc.) by name
where relevant but do NOT invent specific URLs or article titles.

All content reflects your knowledge of publicly known industry developments;
you do not speculate about RBC's internal performance or confidential data.
"""

REPORT_PROMPT_TEMPLATE = """\
Produce a structured monthly executive briefing for the RBC Investor Services
Senior Management Team covering **{prev_month}**.

Today's date is {today}.

---

The report must contain exactly the following numbered sections, each with
3-6 bullet points of substance (no padding):

1. EXECUTIVE SUMMARY
   A tight 4-5 sentence paragraph (not bullets) giving the overall narrative
   for the month.  Lead with the most material development.

2. MARKET & MACRO ENVIRONMENT
   Key equity, fixed income, FX, and rates moves that matter to custodians,
   fund administrators, and asset servicers.

3. REGULATORY & POLICY DEVELOPMENTS
   Significant regulatory updates (SEC, FCA, ESMA, CSSF, OSC, etc.)
   affecting investor services — reporting rules, capital requirements,
   AML/KYC, digital assets regulation, settlement reform.

4. INDUSTRY & COMPETITIVE LANDSCAPE
   Notable M&A, partnerships, product launches, or strategic shifts among
   major custodians and fund servicers (BNY Mellon, State Street, Northern
   Trust, Citi, HSBC, etc.) and FinTech challengers.

5. RISK DASHBOARD
   The top 5 risks for investor services firms for the month.  For each:
   • Risk title (e.g. "Counterparty Credit – EM Sovereign Exposure")
   • Risk category: [MARKET | CREDIT | OPERATIONAL | REGULATORY | CYBER | LIQUIDITY]
   • Brief description (2 sentences max)
   • Direction: ↑ INCREASING | → STABLE | ↓ DECREASING

6. TECHNOLOGY & DIGITAL ASSETS
   AI adoption in fund admin, tokenisation of funds/securities, DLT
   infrastructure, T+1/T+0 settlement progress, and cyber threats relevant
   to securities services.

7. KEY UPCOMING EVENTS — {next_month}
   List 6-8 significant scheduled events for the coming month: regulatory
   deadlines, central bank meetings, major industry conferences, and
   important economic data releases relevant to investor services.

8. SMT WATCH LIST
   3 themes requiring active monitoring by senior management in the next
   30 days, with a single recommended action for each.

---

Format rules:
- Use the section numbers and titles exactly as shown above.
- Within sections 2–6, use bullet points starting with "•".
- For section 7, use a dated list where dates are known: "dd Mmm — [Event]".
- Do not add sub-headings inside sections beyond what is specified.
- Do not include the system prompt in the output.
- Do not add a closing sign-off or footer note.
- Do NOT use any markdown formatting: no **, *, __, _, #, or backticks anywhere in the output.
"""


def generate_report_content(prev_month_year: int, prev_month_month: int) -> str:
    """Call Claude to generate the report text and return it."""
    client = anthropic.Anthropic()

    prev_m = month_label(prev_month_year, prev_month_month)
    next_m = next_month_label(prev_month_year, prev_month_month)
    today_str = date.today().strftime("%d %B %Y")

    prompt = REPORT_PROMPT_TEMPLATE.format(
        prev_month=prev_m,
        next_month=next_m,
        today=today_str,
    )

    print(f"Generating report for {prev_m} via Claude…")
    response = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text


# ── Word document builder ──────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color: str) -> None:
    """Set table cell background colour (OOXML helper)."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def _add_cover_page(doc: Document, prev_month: str, today: date) -> None:
    """Insert a branded cover page."""
    # Top colour band
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    cell = table.cell(0, 0)
    _set_cell_bg(cell, "00429B")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("RBC INVESTOR SERVICES")
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    run.font.name = "Calibri"
    table.rows[0].height = Cm(2.5)

    doc.add_paragraph()

    # Title block
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run("MONTHLY EXECUTIVE REPORT")
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RBC_BLUE
    run.font.name = "Calibri"

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub_p.add_run(f"For the period: {prev_month}")
    run.font.size = Pt(16)
    run.font.color.rgb = RBC_DARK
    run.font.name = "Calibri"

    doc.add_paragraph()

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = date_p.add_run(
        f"Prepared for: Senior Management Team (SMT)\n"
        f"Published: {today.strftime('%d %B %Y')}\n"
        f"Classification: RESTRICTED – SMT ONLY"
    )
    run.font.size = Pt(11)
    run.font.color.rgb = RBC_DARK
    run.font.name = "Calibri"

    # Gold divider
    table2 = doc.add_table(rows=1, cols=1)
    table2.style = "Table Grid"
    cell2 = table2.cell(0, 0)
    _set_cell_bg(cell2, "FFBE00")
    table2.rows[0].height = Cm(0.4)
    table2.cell(0, 0).paragraphs[0].text = ""

    doc.add_page_break()


def _section_heading(doc: Document, title: str) -> None:
    """Add a formatted section heading."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.font.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = RBC_BLUE
    run.font.name = "Calibri"

    # Underline via bottom border
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "00429B")
    pBdr.append(bottom)
    pPr.append(pBdr)


def _strip_md(text: str) -> str:
    """Remove markdown formatting characters from a string."""
    # Bold: **text** and __text__
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text, flags=re.DOTALL)
    text = re.sub(r'__(.+?)__', r'\1', text, flags=re.DOTALL)
    # Italic: *text* and _text_
    text = re.sub(r'\*(.+?)\*', r'\1', text, flags=re.DOTALL)
    text = re.sub(r'_(.+?)_', r'\1', text, flags=re.DOTALL)
    # Inline code
    text = re.sub(r'`(.+?)`', r'\1', text)
    # Leading # heading markers
    text = re.sub(r'^#+\s*', '', text)
    # Any remaining stray * or # characters
    text = text.replace('**', '').replace('*', '').replace('#', '')
    return text.strip()


def _parse_and_render_sections(doc: Document, raw_text: str) -> None:
    """
    Parse the structured text from Claude and render each section into the
    document with appropriate formatting.
    """
    # Split on numbered section headers like "1. EXECUTIVE SUMMARY"
    # Allow lowercase letters and digits in title (e.g. month names in section 7)
    section_pattern = re.compile(
        r"^\s*(\d+)\.\s+([A-Z][A-Za-z0-9 &\—–/\-]+(?:\s*[—–\-]\s*[A-Za-z0-9 &]+)?)\s*$",
        re.MULTILINE,
    )
    matches = list(section_pattern.finditer(raw_text))

    if not matches:
        # Fallback: dump raw text
        doc.add_paragraph(raw_text)
        return

    for i, match in enumerate(matches):
        sec_num = match.group(1)
        sec_title = match.group(2).strip()
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(raw_text)
        body = raw_text[start:end].strip()

        _section_heading(doc, f"{sec_num}. {sec_title}")

        # Section 1 is a prose paragraph; the rest are bullet lists
        if sec_num == "1":
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(_strip_md(body))
            run.font.size = Pt(11)
            run.font.name = "Calibri"
        else:
            for line in body.splitlines():
                line = line.strip()
                if not line:
                    continue

                # Detect if line is a bullet (•, -, or leading *)
                is_bullet = line.startswith("•") or line.startswith("-") or (
                    line.startswith("*") and not line.startswith("**")
                )

                # Detect risk rows with direction arrows (section 5)
                if sec_num == "5" and is_bullet:
                    direction_color = None
                    if "↑ INCREASING" in line:
                        direction_color = RGBColor(0xCC, 0x00, 0x00)
                    elif "↓ DECREASING" in line:
                        direction_color = RGBColor(0x00, 0x7A, 0x33)
                    elif "→ STABLE" in line:
                        direction_color = RGBColor(0xFF, 0x80, 0x00)

                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.space_after = Pt(3)
                    content = _strip_md(re.sub(r'^[•\-\*]\s*', '', line))

                    if direction_color:
                        for marker in ("↑ INCREASING", "→ STABLE", "↓ DECREASING"):
                            if marker in content:
                                pre, _, _ = content.partition(marker)
                                run = p.add_run(pre.rstrip())
                                run.font.size = Pt(10.5)
                                run.font.name = "Calibri"
                                run2 = p.add_run(f"  {marker}")
                                run2.font.size = Pt(10.5)
                                run2.font.bold = True
                                run2.font.color.rgb = direction_color
                                run2.font.name = "Calibri"
                                break
                    else:
                        run = p.add_run(content)
                        run.font.size = Pt(10.5)
                        run.font.name = "Calibri"

                elif is_bullet:
                    p = doc.add_paragraph(style="List Bullet")
                    p.paragraph_format.space_after = Pt(3)
                    run = p.add_run(_strip_md(re.sub(r'^[•\-\*]\s*', '', line)))
                    run.font.size = Pt(10.5)
                    run.font.name = "Calibri"
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.space_after = Pt(3)
                    run = p.add_run(_strip_md(line))
                    run.font.size = Pt(10.5)
                    run.font.name = "Calibri"


def build_word_document(
    report_text: str,
    prev_month_year: int,
    prev_month_month: int,
    today: date,
) -> bytes:
    """
    Build and return the Word document as bytes.
    Uses report_template.docx if present, otherwise builds from scratch.
    """
    if TEMPLATE_FILE.exists():
        doc = Document(str(TEMPLATE_FILE))
        # Clear all existing body paragraphs in the template
        for para in list(doc.paragraphs):
            para._element.getparent().remove(para._element)
        print(f"Using template: {TEMPLATE_FILE}")
    else:
        doc = Document()
        # Page margins
        section = doc.sections[0]
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        print("No template found — building document from scratch.")

    prev_m = month_label(prev_month_year, prev_month_month)
    _add_cover_page(doc, prev_m, today)
    _parse_and_render_sections(doc, report_text)

    # Footer disclaimer
    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run(
        "This document is confidential and intended solely for the RBC Investor "
        "Services Senior Management Team. Do not distribute."
    )
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    run.font.italic = True
    run.font.name = "Calibri"

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Gmail sender ───────────────────────────────────────────────────────────────

def _get_gmail_service():
    """Authenticate and return a Gmail API service object."""
    if not GMAIL_AVAILABLE:
        raise RuntimeError(
            "Gmail API libraries not installed.\n"
            "Run: pip install google-auth-oauthlib google-api-python-client"
        )

    creds = None
    if TOKEN_FILE.exists():
        with open(TOKEN_FILE, "rb") as fh:
            creds = pickle.load(fh)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"\nGmail credentials not found at {CREDENTIALS_FILE}.\n\n"
                    "Setup:\n"
                    "  1. https://console.cloud.google.com/ → enable Gmail API\n"
                    "  2. Create OAuth 2.0 credentials (Desktop App)\n"
                    "  3. Download JSON → save as credentials.json here\n"
                    "  4. Run: python generate_report.py --dry-run  (one-time browser auth)\n"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_FILE), GMAIL_SCOPES
            )
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, "wb") as fh:
            pickle.dump(creds, fh)

    return _gs_build("gmail", "v1", credentials=creds)


def send_email(
    docx_bytes: bytes,
    prev_month_year: int,
    prev_month_month: int,
    today: date,
) -> None:
    """Send the report as an email attachment via Gmail."""
    service = _get_gmail_service()

    prev_m = month_label(prev_month_year, prev_month_month)
    subject = f"RBC Investor Services — Monthly Executive Report: {prev_m}"
    filename = (
        f"RBC_IS_Executive_Report_{prev_month_year}"
        f"_{prev_month_month:02d}.docx"
    )

    body_html = f"""\
<html><body style="font-family:Calibri,Arial,sans-serif;font-size:13px;color:#1a1a2e;">
<p>Dear Ian,</p>

<p>Please find attached the <strong>RBC Investor Services Monthly Executive Report</strong>
for <strong>{prev_m}</strong>, prepared for the Senior Management Team.</p>

<p>The report covers:</p>
<ul>
  <li>Market &amp; Macro Environment</li>
  <li>Regulatory &amp; Policy Developments</li>
  <li>Industry &amp; Competitive Landscape</li>
  <li>Risk Dashboard</li>
  <li>Technology &amp; Digital Assets</li>
  <li>Key Upcoming Events</li>
  <li>SMT Watch List</li>
</ul>

<p>Published: {today.strftime('%d %B %Y')}<br>
Classification: RESTRICTED – SMT ONLY</p>

<hr style="border:1px solid #00429B; width:60%; text-align:left; margin-left:0;">
<p style="font-size:11px;color:#888;">
This message and its attachments are confidential and intended solely for the
named recipient(s). If you have received this in error, please notify the sender
and delete it immediately.
</p>
</body></html>
"""

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["To"] = ", ".join(RECIPIENTS)

    # HTML body
    msg.attach(MIMEText(body_html, "html"))

    # .docx attachment
    part = MIMEBase(
        "application",
        "vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    part.set_payload(docx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"Email sent to: {', '.join(RECIPIENTS)}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate and email the RBC IS Monthly Executive Report",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Save the report locally but do not send the email",
    )
    parser.add_argument(
        "--year",
        type=int,
        default=None,
        help="Override the report year (default: previous month)",
    )
    parser.add_argument(
        "--month",
        type=int,
        default=None,
        help="Override the report month 1–12 (default: previous month)",
    )
    parser.add_argument(
        "--output-dir",
        default=".",
        help="Directory to save the .docx file (default: current directory)",
    )
    args = parser.parse_args()

    today = date.today()

    if args.year and args.month:
        prev_year, prev_month = args.year, args.month
    else:
        prev_year, prev_month = prev_month_info(today)

    # 1. Generate content
    report_text = generate_report_content(prev_year, prev_month)

    # 2. Build Word document
    docx_bytes = build_word_document(report_text, prev_year, prev_month, today)

    # 3. Save locally
    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    filename = f"RBC_IS_Executive_Report_{prev_year}_{prev_month:02d}.docx"
    out_path = out_dir / filename
    out_path.write_bytes(docx_bytes)
    print(f"Report saved: {out_path}")

    # 4. Send email (unless dry-run)
    if args.dry_run:
        print("[Dry run] Email not sent.")
    else:
        send_email(docx_bytes, prev_year, prev_month, today)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
