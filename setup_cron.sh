#!/usr/bin/env bash
# ============================================================================
# RBC Investor Services — Monthly Report Cron Setup
# ============================================================================
# Installs a cron job that runs generate_report.py on the 1st of every month
# at 07:00 local time.
#
# Usage:
#   chmod +x setup_cron.sh
#   ./setup_cron.sh
#
# To remove the cron job later:
#   crontab -l | grep -v 'generate_report' | crontab -
# ============================================================================

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN="$(command -v python3)"
REPORT_SCRIPT="${SCRIPT_DIR}/generate_report.py"
LOG_FILE="${SCRIPT_DIR}/report_cron.log"

echo "Setting up monthly cron job..."
echo "  Script : ${REPORT_SCRIPT}"
echo "  Python : ${PYTHON_BIN}"
echo "  Log    : ${LOG_FILE}"
echo ""

# Cron expression: minute=0 hour=7 day=1 month=* weekday=*
CRON_LINE="0 7 1 * * cd \"${SCRIPT_DIR}\" && \"${PYTHON_BIN}\" \"${REPORT_SCRIPT}\" >> \"${LOG_FILE}\" 2>&1"

# Add only if not already present
( crontab -l 2>/dev/null | grep -F 'generate_report'; true ) | grep -q 'generate_report' && {
    echo "Cron job already exists — no changes made."
    exit 0
}

( crontab -l 2>/dev/null; echo "${CRON_LINE}" ) | crontab -

echo "Cron job installed. Current crontab:"
crontab -l | grep 'generate_report'
echo ""
echo "The report will run at 07:00 on the 1st of each month."
echo ""
echo "Prerequisites before the first run:"
echo "  1. pip install -r requirements.txt"
echo "  2. Place credentials.json (Gmail OAuth2) in: ${SCRIPT_DIR}"
echo "  3. Authenticate once interactively:"
echo "       python3 generate_report.py --dry-run"
echo "     (a browser window will open for Gmail authorisation)"
echo "  4. Optionally place report_template.docx in: ${SCRIPT_DIR}"
echo "     (if omitted, a built-in RBC-branded template is used)"
