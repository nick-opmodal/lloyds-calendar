#!/bin/bash
# Regenerate the Lloyds Deadline Calendar ICS feed.
# Called weekly by cron (Mondays 06:00).

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$SCRIPT_DIR/.venv/bin/python"

cd "$SCRIPT_DIR"

echo "=== Lloyds Deadline Calendar Regeneration — $(date '+%Y-%m-%d %H:%M') ==="

# Generate ICS files
echo "→ Generating ICS feeds..."
$PYTHON generate_lloyds_calendar.py --weeks 12 -o lloyds-deadlines.ics
$PYTHON generate_lloyds_calendar.py --weeks 12 --curated-only -o lloyds-deadlines-curated.ics

# Git commit if there are changes
if ! git diff --quiet lloyds-deadlines.ics lloyds-deadlines-curated.ics 2>/dev/null; then
    git add lloyds-deadlines.ics lloyds-deadlines-curated.ics
    git commit -m "Weekly regeneration — $(date '+%Y-%m-%d')" 2>/dev/null || true
    echo "→ Changes committed."
else
    echo "→ No changes to commit."
fi

# Push to GitHub Pages (the feed is served from main at calendar.subscriptionmarket.news).
# Self-healing: pushes whenever local is ahead of origin, not just on regeneration days.
if [ -n "$(git rev-list origin/main..HEAD 2>/dev/null)" ]; then
    git pull --rebase 2>/dev/null || true
    git push origin main
    echo "→ Pushed to GitHub Pages: $(git rev-parse --short HEAD)"
else
    echo "→ Nothing to push (origin up to date)."
fi

echo "Done."
