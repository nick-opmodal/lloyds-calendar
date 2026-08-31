#!/usr/bin/env python3
"""
Lloyd's Business Timetable -> annotated ICS calendar feed.

Weekly pipeline for the Opmodal 'Lloyd's Deadlines' subscribable calendar:

  1. Download the official XLSX export from lloyds.com
  2. Parse deadlines (header row and columns auto-detected)
  3. Filter to a rolling window and optional category/participant filters
  4. Attach editorial commentary from commentary.yaml (keyword rules)
  5. Emit an RFC 5545 .ics with Europe/London timezone and stable UIDs
     (stable UIDs mean subscribers' calendars update in place, no duplicates)

Usage:
  python generate_lloyds_calendar.py                    # normal weekly run
  python generate_lloyds_calendar.py --inspect          # show detected columns + sample rows, then exit
  python generate_lloyds_calendar.py --curated-only     # only events that matched a commentary rule
  python generate_lloyds_calendar.py --xlsx local.xlsx  # parse a local file instead of downloading
  python generate_lloyds_calendar.py --weeks 12 -o out.ics

Dependencies:  pip install requests openpyxl pyyaml icalendar
(icalendar is optional - used only for output validation)

Deploy: publish the output at a stable URL, e.g. https://opmodal.com/lloyds-deadlines.ics
Subscribers add it once ("Subscribe from web" in Outlook / "New Calendar Subscription"
in Apple Calendar) and receive every weekly regeneration automatically.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import re
import sys
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

XLSX_URL = "https://www.lloyds.com/market-resources/business-timetable/xlsx"
XLS_URL = "https://www.lloyds.com/market-resources/business-timetable/xls"
TIMETABLE_URL = "https://www.lloyds.com/market-resources/business-timetable"
CAL_NAME = "Lloyd's Deadlines - by Opmodal"
CAL_DESC = (
    "Key Lloyd's business timetable deadlines with commentary from the "
    "Submission Markets team at Opmodal. Data: lloyds.com. "
    "Always verify against the official timetable."
)
UID_DOMAIN = "opmodal.com"
DEFAULT_WEEKS_AHEAD = 8

# Column auto-detection: header cell must contain one of these substrings
# (case-insensitive). First match wins. Adjust after an --inspect run if
# Lloyd's ever renames a column.
# Verified against live XLSX export (2026-07-28):
#   Row 9 headers: Current Deadline, Deadline Time, Original Deadline, Title,
#   Market Participant, Submission Link, Submission Instructions, ...
COLUMN_ALIASES = {
    "date":            ["current deadline", "deadline date", "due date"],
    "title":           ["title", "name", "event"],
    "description":     ["submission instructions"],
    "additional_info": ["additionalinformation", "additional information"],
    "category":        ["categories"],
    "participant":     ["market participant", "particpant", "participant", "audience"],
    "region":          ["region"],
    "time":            ["deadline time", "time due", "time"],
}

# Optional row filters. Leave empty to include everything.
INCLUDE_PARTICIPANTS: list[str] = []
EXCLUDE_CATEGORIES: list[str] = []


# --------------------------------------------------------------------------- #
#  Fetch & parse                                                              #
# --------------------------------------------------------------------------- #

def download_timetable() -> bytes:
    """Download the timetable export. Lloyd's changed the export endpoint from
    .xlsx to a legacy .xls (NPOI-generated OLE2) in Aug 2026; try xlsx first,
    fall back to xls, retry transient failures (the site 504s under load)."""
    import time

    import requests

    urls = (XLSX_URL, XLS_URL)
    last_err: Exception | None = None
    for attempt in range(3):
        for url in urls:
            try:
                r = requests.get(url, timeout=60,
                                 headers={"User-Agent": "OpmodalCalendarBot/1.0"})
                r.raise_for_status()
                content = r.content
                if content[:2] == b"PK" or content[:4] == b"\xd0\xcf\x11\xe0":
                    return content
                last_err = RuntimeError(
                    f"unexpected file format from {url} ({len(content)} bytes)"
                )
            except Exception as e:  # noqa: BLE001 - retry whatever the site throws
                last_err = e
        time.sleep(10 * (attempt + 1))
    raise RuntimeError(f"could not download Lloyd's timetable export: {last_err}")


def detect_columns(header_cells: list) -> dict:
    """Map logical field names to column indexes using COLUMN_ALIASES."""
    mapping = {}
    lowered = [(i, str(c).strip().lower()) for i, c in enumerate(header_cells) if c]
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            hit = next((i for i, h in lowered if alias in h and i not in mapping.values()), None)
            if hit is not None:
                mapping[field] = hit
                break
    return mapping


def parse_rows(data_bytes: bytes, inspect: bool = False) -> list[dict]:
    if data_bytes[:2] == b"PK":
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data_bytes), read_only=True, data_only=True)
        ws = wb.active
        rows = [[c for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    else:
        # Legacy .xls (OLE2) — NPOI export. Calamine reads both formats.
        from python_calamine import CalamineWorkbook
        wb = CalamineWorkbook.from_filelike(io.BytesIO(data_bytes))
        rows = wb.get_sheet_by_index(0).to_python()

    # Find the header row: first row whose cells match a 'date' alias AND a 'title' alias.
    header_idx, mapping = None, {}
    for i, row in enumerate(rows):
        m = detect_columns(list(row))
        if "date" in m and "title" in m:
            header_idx, mapping = i, m
            break
    if header_idx is None:
        raise RuntimeError(
            "Could not locate a header row in the XLSX. Run with --inspect "
            "and adjust COLUMN_ALIASES to match the current export layout."
        )

    if inspect:
        print(f"Header row: {header_idx + 1}")
        print(f"Raw headers: {rows[header_idx]}")
        print(f"Detected mapping: { {k: rows[header_idx][v] for k, v in mapping.items()} }")
        for r in rows[header_idx + 1: header_idx + 6]:
            print("  sample:", r)
        sys.exit(0)

    events = []
    for row in rows[header_idx + 1:]:
        if not row or all(c in (None, "") for c in row):
            continue
        get = lambda f: (str(row[mapping[f]]).strip()
                         if f in mapping and mapping[f] < len(row) and row[mapping[f]] is not None
                         else "")
        d = coerce_date(row[mapping["date"]] if mapping["date"] < len(row) else None)
        title = get("title")
        if not d or not title:
            continue
        desc = get("description")
        addl = get("additional_info")
        # Fall back to AdditionalInformation when Submission Instructions is empty
        if not desc:
            desc = addl
        events.append({
            "date": d,
            "time": coerce_time(get("time")),
            "title": title,
            "description": desc,
            "category": get("category"),
            "participant": get("participant"),
            "region": get("region"),
        })
    return events


def coerce_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if v is None:
        return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d %B %Y", "%a %d %b %Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def coerce_time(s: str):
    """'13:00 BST' / '1pm' / '14:00' -> 'HHMM' string, else None (all-day)."""
    if not s:
        return None
    m = re.search(r"(\d{1,2})[:.](\d{2})", s)
    if m:
        return f"{int(m.group(1)):02d}{m.group(2)}"
    m = re.search(r"(\d{1,2})\s*(am|pm)", s, re.I)
    if m:
        h = int(m.group(1)) % 12 + (12 if m.group(2).lower() == "pm" else 0)
        return f"{h:02d}00"
    return None


# --------------------------------------------------------------------------- #
#  Editorial layer                                                            #
# --------------------------------------------------------------------------- #

def load_commentary(path: Path) -> list[dict]:
    import yaml
    if not path.exists():
        print(f"note: no commentary file at {path}; events will carry Lloyd's text only")
        return []
    rules = yaml.safe_load(path.read_text(encoding="utf-8")) or []
    for r in rules:
        r["_re"] = re.compile(r["match"], re.I)
    return rules


def annotate(event: dict, rules: list[dict]) -> str | None:
    """Return commentary for the first matching rule, or None."""
    haystack = f"{event['title']} {event['description']} {event['category']}"
    for r in rules:
        if r["_re"].search(haystack):
            if "month" in r and event["date"].strftime("%Y-%m") != r["month"]:
                continue
            return r["comment"]
    return None


# --------------------------------------------------------------------------- #
#  ICS generation                                                             #
# --------------------------------------------------------------------------- #

def esc(s: str) -> str:
    return (s.replace("\\", "\\\\").replace(";", "\\;")
             .replace(",", "\\,").replace("\n", "\\n"))


def fold(line: str) -> str:
    out, b = [], line.encode("utf-8")
    while len(b) > 73:
        cut = 73
        while cut > 0 and (b[cut] & 0xC0) == 0x80:  # don't split a UTF-8 char
            cut -= 1
        out.append(b[:cut].decode("utf-8"))
        b = b" " + b[cut:]
    out.append(b.decode("utf-8"))
    return "\r\n".join(out)


VTIMEZONE = [
    "BEGIN:VTIMEZONE", "TZID:Europe/London",
    "BEGIN:DAYLIGHT", "TZOFFSETFROM:+0000", "TZOFFSETTO:+0100", "TZNAME:BST",
    "DTSTART:19700329T010000", "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU", "END:DAYLIGHT",
    "BEGIN:STANDARD", "TZOFFSETFROM:+0100", "TZOFFSETTO:+0000", "TZNAME:GMT",
    "DTSTART:19701025T020000", "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU", "END:STANDARD",
    "END:VTIMEZONE",
]


def build_ics(events: list[dict], rules: list[dict], curated_only: bool) -> str:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        "PRODID:-//Opmodal//Lloyds Business Timetable Feed 1.0//EN",
        "CALSCALE:GREGORIAN", "METHOD:PUBLISH",
        f"X-WR-CALNAME:{esc(CAL_NAME)}",
        f"X-WR-CALDESC:{esc(CAL_DESC)}",
        "X-WR-TIMEZONE:Europe/London",
        "X-PUBLISHED-TTL:PT12H",          # hint: refresh subscribers twice daily
        *VTIMEZONE,
    ]
    n = 0
    # Pre-compute distinguishing subtitles for events with duplicate titles on same date
    from collections import defaultdict
    by_date_title = defaultdict(list)
    for ev in events:
        by_date_title[(ev["date"], ev["title"])].append(ev)
    
    def _common_prefix(a: str, b: str) -> int:
        """Length of longest common prefix between two strings."""
        i = 0
        for ca, cb in zip(a, b):
            if ca != cb:
                break
            i += 1
        return i
    
    def _subtitle(ev: dict) -> str:
        """Return a short distinguishing subtitle from the description."""
        raw_desc = ev.get("description") or ""
        if not raw_desc:
            return ""
        group = by_date_title[(ev["date"], ev["title"])]
        if len(group) == 1:
            # Singleton: use first 70 chars, stripping title prefix
            excerpt = raw_desc
            if excerpt.lower().startswith(ev["title"].lower()):
                excerpt = excerpt[len(ev["title"]):].strip().lstrip("-–— ").strip()
            return excerpt[:70].strip()
        # Duplicate titles: find the longest shared prefix with any peer,
        # then use the distinguishing suffix beyond that
        descs = [e.get("description") or "" for e in group]
        my_idx = group.index(ev)
        longest_shared = 0
        for j, other in enumerate(descs):
            if j == my_idx:
                continue
            longest_shared = max(longest_shared, _common_prefix(raw_desc, other))
        # Walk back to a clean cut point
        while longest_shared > 0 and raw_desc[longest_shared-1:longest_shared] not in (" ", "-", "–", "—", ","):
            longest_shared -= 1
        suffix = raw_desc[longest_shared:].strip().lstrip("-–— ").strip()
        # Strip leading non-word characters (bullets, tabs, etc.)
        suffix = re.sub(r'^[^\w]+', '', suffix).strip()
        return suffix[:70].strip()
    
    for ev in sorted(events, key=lambda e: (e["date"], e["title"])):
        comment = annotate(ev, rules)
        if curated_only and comment is None:
            continue
        n += 1
        dt = ev["date"].strftime("%Y%m%d")
        # Stable UID: same date+title+description -> same UID on every run -> clean updates.
        uid = hashlib.sha1(f"{dt}|{ev['title']}|{ev['description']}".encode()).hexdigest()[:16]
        lines += ["BEGIN:VEVENT",
                  f"UID:lloyds-{dt}-{uid}@{UID_DOMAIN}",
                  f"DTSTAMP:{stamp}"]
        if ev["time"]:
            lines.append(f"DTSTART;TZID=Europe/London:{dt}T{ev['time']}00")
        else:
            lines.append(f"DTSTART;VALUE=DATE:{dt}")
        # Build a distinctive summary using the smart subtitle extractor
        subtitle = _subtitle(ev)
        if subtitle:
            summary = f"Lloyd's: {ev['title']} — {subtitle}"
        else:
            summary = f"Lloyd's: {ev['title']}"
        lines.append(f"SUMMARY:{esc(summary)}")
        desc = ev["description"] or ev["title"]
        if comment:
            desc += f"\n\n▸ WHY IT MATTERS (Submission Markets): {comment}"
        if ev["participant"]:
            desc += f"\n\nApplies to: {ev['participant']}"
        desc += f"\n\nOfficial entry: {TIMETABLE_URL}"
        lines.append(f"DESCRIPTION:{esc(desc)}")
        lines.append(f"URL:{TIMETABLE_URL}")
        if ev["category"]:
            lines.append(f"CATEGORIES:{esc(ev['category'])}")
        lines += ["BEGIN:VALARM", "ACTION:DISPLAY",
                  f"DESCRIPTION:{esc('Lloyd’s deadline tomorrow: ' + ev['title'])}",
                  "TRIGGER:-P1D", "END:VALARM", "END:VEVENT"]
    lines.append("END:VCALENDAR")
    print(f"{n} events written")
    return "\r\n".join(fold(l) for l in lines) + "\r\n"


def validate(ics_text: str) -> None:
    try:
        from icalendar import Calendar
    except ImportError:
        print("note: 'icalendar' not installed, skipping validation")
        return
    cal = Calendar.from_ical(ics_text.encode("utf-8"))
    count = sum(1 for c in cal.walk("VEVENT"))
    print(f"validation OK: {count} events parse cleanly")


# --------------------------------------------------------------------------- #
#  Main                                                                       #
# --------------------------------------------------------------------------- #

def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    ap.add_argument("-o", "--output", default="lloyds-deadlines.ics")
    ap.add_argument("--commentary", default="commentary.yaml")
    ap.add_argument("--weeks", type=int, default=DEFAULT_WEEKS_AHEAD,
                    help=f"rolling window in weeks (default {DEFAULT_WEEKS_AHEAD})")
    ap.add_argument("--xlsx", help="parse a local XLSX instead of downloading")
    ap.add_argument("--inspect", action="store_true",
                    help="print detected header/columns and sample rows, then exit")
    ap.add_argument("--curated-only", action="store_true",
                    help="include only events that matched a commentary rule")
    args = ap.parse_args()

    xlsx = Path(args.xlsx).read_bytes() if args.xlsx else download_timetable()
    events = parse_rows(xlsx, inspect=args.inspect)

    today = date.today()
    horizon = today + timedelta(weeks=args.weeks)
    events = [e for e in events if today <= e["date"] <= horizon]
    if INCLUDE_PARTICIPANTS:
        events = [e for e in events
                  if any(p in e["participant"].lower() for p in INCLUDE_PARTICIPANTS)]
    if EXCLUDE_CATEGORIES:
        events = [e for e in events
                  if not any(c in e["category"].lower() for c in EXCLUDE_CATEGORIES)]
    print(f"{len(events)} events in window {today} -> {horizon}")

    rules = load_commentary(Path(args.commentary))
    ics = build_ics(events, rules, args.curated_only)
    validate(ics)
    Path(args.output).write_bytes(ics.encode("utf-8"))
    print(f"wrote {args.output}")


if __name__ == "__main__":
    main()
