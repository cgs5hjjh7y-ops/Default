#!/usr/bin/env python3
"""
RBC Investor Services — Monthly Executive Report Generator
============================================================
Generates an AI-authored executive report for the RBC Investor Services
Senior Management Team (SMT) covering the preceding calendar month.

The report is produced as a Word (.docx) document and emailed via Gmail
SMTP (App Password) to the configured recipients.

Usage:
  python generate_report.py               # Generate & email for the previous month
  python generate_report.py --dry-run     # Generate & save locally, do not email
  python generate_report.py --year 2026 --month 2   # Target a specific month

Scheduling:
  Run via cron on the 1st of each month (see setup_cron.sh).

Gmail App Password setup:
  1. Enable 2-Step Verification on your Google account
  2. Visit https://myaccount.google.com/apppasswords
  3. Generate a 16-character App Password
  4. Set environment variables:
       GMAIL_ADDRESS=iansinclair3011@gmail.com
       GMAIL_APP_PASSWORD=<your-16-char-password>

Word template:
  If report_template.docx exists in this directory it will be used as the
  base document (styles, header/footer, logo, etc.).  Otherwise a clean
  document is built programmatically from the RBC brand colours.
"""

import argparse
import io
import os
import re
import smtplib
import ssl
import sys
from calendar import month_name
from datetime import date, datetime
from email.mime.application import MIMEApplication
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
    "paul.p.burd@rbc.com",
    "joel.kornblum@rbc.com",
    "christine.knott@rbc.com",
]

TEMPLATE_FILE = Path("report_template.docx")

# RBC brand colours
RBC_BLUE = RGBColor(0x00, 0x42, 0x9B)   # #00429B
RBC_GOLD = RGBColor(0xFF, 0xBE, 0x00)   # #FFBE00
RBC_DARK = RGBColor(0x1A, 0x1A, 0x2E)   # near-black

# Gmail SMTP — set these as environment variables (or GitHub Secrets)
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "iansinclair3011@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")


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
   Output ONLY a pipe-separated table of the top 5 risks for investor services
   firms this month.
   Columns: Risk Title | Category | Description | Direction
   Category values: MARKET | CREDIT | OPERATIONAL | REGULATORY | CYBER | LIQUIDITY
   Direction values: ↑ INCREASING | → STABLE | ↓ DECREASING
   Description: 1-2 sentences maximum per risk.
   Do NOT include a separator line.  Do NOT include any text other than the
   header row and the 5 data rows.

6. TECHNOLOGY & DIGITAL ASSETS
   AI adoption in fund admin, tokenisation of funds/securities, DLT
   infrastructure, T+1/T+0 settlement progress, and cyber threats relevant
   to securities services.

7. KEY UPCOMING EVENTS — {next_month}
   List 6-8 significant scheduled events for the coming month: regulatory
   deadlines, central bank meetings, major industry conferences, and
   important economic data releases relevant to investor services.

8. SMT WATCH LIST
   3 themes that Canadian-based institutional clients of RBC Investor Services
   (pension funds, asset managers, insurance companies, sovereign wealth funds,
   and fund companies domiciled or operating in Canada) should be actively
   monitoring over the next 30 days.  For each theme, address:
   • The global development or risk driving the theme (e.g. tariffs, rate
     moves, regulatory change, geopolitical shift, FX volatility)
   • The specific impact on Canadian clients' assets, operations, or
     reporting obligations held or serviced through RBC Investor Services
   • A single, concrete recommended action for RBC IS senior management to
     take proactively on behalf of, or in anticipation of, client need

9. TECHNOLOGY VENDOR LANDSCAPE
   Key developments from the major technology vendors serving the investor
   services industry: SimCorp, SS&C Technologies, Broadridge, FIS Global,
   Temenos, Charles River (MSCI), and notable FinTech challengers (e.g.
   Nasdaq Financial Technology, Clearstream, SWIFT, Taskize).  Cover product
   launches, partnerships, M&A activity, and strategic positioning relevant
   to custody, fund administration, and data/reporting services.

10. KEY MACRO ECONOMIC INDICES
   Output ONLY a pipe-separated table of the following indices as at end of
   {prev_month}.  Use your best knowledge of approximate closing levels; do
   not invent fictitious precision — round figures are acceptable.
   Columns: Index | Level | Month Change | YTD Change
   Rows (in this order): S&P 500, FTSE 100, EURO STOXX 50, Nikkei 225,
   EUR/USD, GBP/USD, USD/JPY, 10Y UST Yield, 10Y Bund Yield, Brent Crude,
   Gold, VIX.
   Do NOT include a separator line (no --- rows).  Do NOT include any text
   other than the header row and the 12 data rows.

11. CUSTODIAN & PEER INSTITUTION NEWS
   Output ONLY a pipe-separated table providing the most significant recent
   development for each of the following institutions during {prev_month}.
   Columns: Institution | Key Development | Significance
   If there is no material news for an institution this month, write
   "No material news this month" in the Key Development column and leave
   Significance as "—".
   Include EVERY institution below as its own row — do not omit any.
   Institutions (in this order):
   BNY (Bank of New York Mellon), State Street Corporation,
   JPMorgan Chase & Co., Citigroup Inc. (Citi), HSBC Holdings plc,
   BNP Paribas, Northern Trust Corporation, Deutsche Bank AG, UBS Group AG,
   RBC Investor Services, Societe Generale Securities Services,
   Standard Chartered plc, Brown Brothers Harriman & Co.,
   Clearstream, Euroclear, CIBC Mellon,
   Coinbase Custody (Coinbase Prime), BitGo, Fidelity Digital Assets,
   Anchorage Digital, Fireblocks, Copper, Komainu, Zodia Custody,
   NYDIG, Taurus, Kraken.
   Do NOT include a separator line.  Do NOT include any text other than the
   header row and the 27 data rows.

---

Format rules:
- Use the section numbers and titles exactly as shown above.
- Within sections 2–4, 6, 8, and 9, use bullet points starting with "•".
- For section 7, use a dated list where dates are known: "dd Mmm — [Event]".
- For sections 5, 10, and 11, output only pipe-separated rows as instructed — no bullets.
- Do not add sub-headings inside sections beyond what is specified.
- Do not include the system prompt in the output.
- Do not add a closing sign-off or footer note.
- Do NOT use any markdown formatting: no **, *, __, _, #, or backticks anywhere in the output.
- In company and index names always use "&" (the ampersand symbol), never spell it out as "and". Examples: S&P 500, SS&C Technologies, Brown Brothers Harriman & Co., JPMorgan Chase & Co., RBC Investor Services, Societe Generale Securities Services.
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
        max_tokens=8192,
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


def _fill_template_header(doc: Document, today: date) -> None:
    """Populate the template's memo-style header table with report metadata."""
    if not doc.tables:
        return

    def _set_cell(cell, text: str, bold: bool = False) -> None:
        para = cell.paragraphs[0]
        for run in list(para.runs):
            run._r.getparent().remove(run._r)
        r = para.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(16) if bold else Pt(11)
        r.font.bold = bold

    table = doc.tables[0]
    # Row 1: Title (merged across columns) — bold, larger font
    _set_cell(table.cell(1, 0), "Monthly Executive Report", bold=True)
    table.cell(1, 0).paragraphs[0].paragraph_format.space_after = Pt(6)
    # Row 2: To
    _set_cell(table.cell(2, 1), "Senior Management Team")
    # Row 3: From
    _set_cell(table.cell(3, 1), "Ian Sinclair")
    # Row 4: Date
    _set_cell(table.cell(4, 1), today.strftime("%d %B %Y"))


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


def _render_pipe_table(doc: Document, body: str) -> None:
    """Render pipe-separated lines from *body* as a styled Word table."""
    lines = [l.strip() for l in body.splitlines() if l.strip() and "|" in l]
    if not lines:
        return

    rows_data = [[c.strip() for c in line.split("|")] for line in lines]
    # Determine column count from the widest row
    n_cols = max(len(r) for r in rows_data)

    table = doc.add_table(rows=0, cols=n_cols)
    table.style = "Table Grid"

    for row_idx, row_data in enumerate(rows_data):
        row = table.add_row()
        is_header = row_idx == 0
        for col_idx in range(n_cols):
            cell = row.cells[col_idx]
            text = row_data[col_idx] if col_idx < len(row_data) else ""
            if is_header:
                _set_cell_bg(cell, "00429B")
            elif row_idx % 2 == 0:
                _set_cell_bg(cell, "EEF3FA")  # light blue stripe
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.space_before = Pt(2)
            run = para.add_run(text)
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            run.font.bold = is_header
            if is_header:
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            else:
                # Direction arrow colouring (risk table)
                if "↑ INCREASING" in text:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                elif "→ STABLE" in text:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0x80, 0x00)
                elif "↓ DECREASING" in text:
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(0x00, 0x7A, 0x33)
                # Positive/negative change colouring (macro indices table)
                elif col_idx > 0 and text.startswith("+"):
                    run.font.color.rgb = RGBColor(0x00, 0x7A, 0x33)
                elif col_idx > 0 and text.startswith("-"):
                    run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)


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

        # Sections rendered as pipe-delimited tables
        if sec_num in ("5", "10", "11"):
            _render_pipe_table(doc, body)
            continue

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

                if is_bullet:
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
        # Remove top-level paragraphs but keep the template tables
        for para in list(doc.paragraphs):
            para._element.getparent().remove(para._element)
        # Fill in the memo header table from the template
        _fill_template_header(doc, today)
        # Remove the Summary placeholder table (index 1) if present
        if len(doc.tables) > 1:
            tbl = doc.tables[1]
            tbl._element.getparent().remove(tbl._element)
        print(f"Using template: {TEMPLATE_FILE}")
    else:
        doc = Document()
        section = doc.sections[0]
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
        prev_m = month_label(prev_month_year, prev_month_month)
        _add_cover_page(doc, prev_m, today)
        print("No template found — building document from scratch.")

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

    # AI-generated disclaimer
    doc.add_paragraph()
    ai_disclaimer_p = doc.add_paragraph()
    ai_disclaimer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ai_run = ai_disclaimer_p.add_run(
        "DISCLAIMER: This report has been generated autonomously by Claude, an AI assistant "
        "developed by Anthropic, using agentic AI techniques and publicly available information. "
        "The content has not been reviewed or verified by a human and may contain errors, "
        "omissions, or inaccuracies. It should not be relied upon as financial, investment, or "
        "professional advice. Readers are encouraged to independently verify any information "
        "before making decisions based on this report."
    )
    ai_run.font.size = Pt(7)
    ai_run.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    ai_run.font.italic = True
    ai_run.font.name = "Calibri"

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ── Gmail SMTP sender ──────────────────────────────────────────────────────────

def send_email(
    docx_bytes: bytes,
    prev_month_year: int,
    prev_month_month: int,
    today: date,
) -> None:
    """Send the report as an email attachment via Gmail SMTP (App Password)."""
    if not GMAIL_APP_PASSWORD:
        raise RuntimeError(
            "GMAIL_APP_PASSWORD environment variable is not set.\n"
            "Set it to your 16-character Gmail App Password."
        )

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
  <li>Technology Vendor Landscape</li>
  <li>Key Macro Economic Indices</li>
  <li>Custodian &amp; Peer Institution News</li>
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
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = ", ".join(RECIPIENTS)

    msg.attach(MIMEText(body_html, "html"))

    attachment = MIMEApplication(docx_bytes, Name=filename)
    attachment["Content-Disposition"] = f'attachment; filename="{filename}"'
    msg.attach(attachment)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        smtp.send_message(msg)

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
