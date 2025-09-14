#!/usr/bin/env python3
"""
Weekly Google Calendar color-hours report (Sun→Sat window).

- Authenticates with Google (Calendar API, read-only).
- Uses LAST Sunday 00:00 → THIS Sunday 00:00 (this Sunday excluded).
- Fetches events (primary by default, or all calendars with --all-calendars).
- Clips to the window; handles all-day and multi-day.
- Totals duration by color ID:
    * If an event has an explicit event colorId (1..11), use that.
    * Otherwise use the calendar's default colorId (1..24).
- Prints a summary and writes a CSV (default: weekly_color_hours.csv).

First-time setup:
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib python-dateutil
Place your OAuth Desktop client JSON as credentials.json next to this script.
"""

from __future__ import annotations

import argparse
import csv
import os
from collections import defaultdict
from datetime import datetime, timedelta, time, date

from dateutil.tz import gettz

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ---------- Config ----------
SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]
DEFAULT_TZ = "America/Chicago"
CSV_FILENAME = "weekly_color_hours.csv"
COUNT_ALL_DAY_AS_24H = True  # if False, all-day events are ignored
# ----------------------------


# ---------- Auth ----------
def auth_service():
    """Return an authenticated Google Calendar service."""
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                raise FileNotFoundError(
                    "Missing credentials.json (OAuth Desktop client). Put it next to this script."
                )
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            try:
                creds = flow.run_local_server(port=0)
            except Exception:
                print("[INFO] Falling back to console auth (copy/paste).")
                creds = flow.run_console()

        with open("token.json", "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return build("calendar", "v3", credentials=creds)
# --------------------------


# ---------- Time window (Sun→Sat) ----------
def last_sun_to_sat(tzname: str):
    """
    Return [start, end) for LAST Sun→Sat in the given tz.
    start = last Sunday's 00:00
    end   = this Sunday's 00:00 (EXCLUSIVE)  → so Saturday night is included, this Sunday is excluded.
    """
    tz = gettz(tzname)
    now = datetime.now(tz)

    # Normalize to today's local midnight
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)

    # Monday=0 … Sunday=6  → days since *last* Sunday
    days_since_last_sun = (today.weekday() + 1) % 7

    # End is the most recent Sunday 00:00 that is <= today (today if it's Sunday)
    end = today - timedelta(days=days_since_last_sun)

    # Start is the Sunday a week earlier
    start = end - timedelta(days=7)

    return start, end



# ---------- Helpers ----------
def clamp_interval(start, end, clamp_start, clamp_end):
    """Return overlap [start,end) ∩ [clamp_start,clamp_end) or None."""
    s = max(start, clamp_start)
    e = min(end, clamp_end)
    return None if s >= e else (s, e)


def parse_event_times(evt, tzname: str):
    """
    Return (start_dt, end_dt, is_all_day), localized to tzname.
    Handles 'dateTime' and all-day 'date'.
    """
    tz = gettz(tzname)
    start = evt.get("start", {})
    end = evt.get("end", {})

    if "dateTime" in start:
        s = datetime.fromisoformat(start["dateTime"].replace("Z", "+00:00")).astimezone(tz)
        e = datetime.fromisoformat(end["dateTime"].replace("Z", "+00:00")).astimezone(tz)
        return s, e, False
    elif "date" in start:
        s_date = date.fromisoformat(start["date"])
        e_date = date.fromisoformat(end["date"])  # exclusive end
        s = datetime.combine(s_date, time(0, 0, 0), tz)
        e = datetime.combine(e_date, time(0, 0, 0), tz)
        return s, e, True
    else:
        # Unexpected; treat as zero-length now
        now = datetime.now(tz)
        return now, now, False


def get_colors_map(service):
    """Return (event_colors, calendar_colors) mapping id -> hex."""
    colors = service.colors().get().execute()
    event_colors = {k: v.get("background") for k, v in colors.get("event", {}).items()}
    calendar_colors = {k: v.get("background") for k, v in colors.get("calendar", {}).items()}
    return event_colors, calendar_colors


def list_calendars(service, include_hidden: bool = False):
    """Return your calendar list (optionally skipping hidden)."""
    cals = []
    page = service.calendarList().list().execute()
    cals += page.get("items", [])
    token = page.get("nextPageToken")
    while token:
        page = service.calendarList().list(pageToken=token).execute()
        cals += page.get("items", [])
        token = page.get("nextPageToken")
    if not include_hidden:
        cals = [c for c in cals if not c.get("hidden")]
    return cals


def fetch_events_for_calendar(service, calendar_id: str, time_min_iso: str, time_max_iso: str):
    """Get all events from a given calendar in [time_min, time_max)."""
    try:
        events = []
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=time_min_iso,
            timeMax=time_max_iso,
            singleEvents=True,
            orderBy="startTime",
            maxResults=2500,
        ).execute()
        events.extend(resp.get("items", []))
        token = resp.get("nextPageToken")
        while token:
            resp = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min_iso,
                timeMax=time_max_iso,
                singleEvents=True,
                orderBy="startTime",
                maxResults=2500,
                pageToken=token,
            ).execute()
            events.extend(resp.get("items", []))
            token = resp.get("nextPageToken")
        return events
    except HttpError as e:
        print(f"[ERROR] Calendar API error ({calendar_id}): {e}")
        return []
    except Exception as e:
        print(f"[ERROR] Unexpected error ({calendar_id}): {e}")
        return []
# ---------------------------


def main():
    parser = argparse.ArgumentParser(description="Weekly Google Calendar color-hours report (Sun→Sat)")
    parser.add_argument("--tz", default=DEFAULT_TZ, help="IANA timezone, e.g. America/Chicago")
    parser.add_argument("--csv", default=CSV_FILENAME, help="Output CSV filename")
    parser.add_argument("--ignore-all-day", action="store_true",
                        help="Ignore all-day events instead of counting them as 24h/day")
    parser.add_argument("--all-calendars", action="store_true",
                        help="Include all visible calendars (default: primary only)")
    args = parser.parse_args()

    tzname = args.tz
    count_all_day = COUNT_ALL_DAY_AS_24H and not args.ignore_all_day

    print("[INFO] Authenticating…")
    service = auth_service()
    print("[INFO] Auth OK")

    # Sun→Sat window: last Sunday 00:00 to this Sunday 00:00 (this Sunday excluded)
    week_start, week_end = last_sun_to_sat(tzname)
    print(f"[INFO] Window: {week_start} → {week_end} ({tzname})")

    time_min_iso = week_start.isoformat()
    time_max_iso = week_end.isoformat()

    # Color tables
    event_colors, calendar_colors = get_colors_map(service)

    # Which calendars to read
    if args.all_calendars:
        calendars = list_calendars(service)
    else:
        calendars = [service.calendarList().get(calendarId="primary").execute()]

    # Sum seconds per color id (event ids 1..11, or calendar default ids 1..24)
    seconds_by_color_id = defaultdict(float)
    events_total = 0

    for cal in calendars:
        cal_id = cal["id"]
        cal_default_color_id = cal.get("colorId")  # this calendar's default color id (1..24)
        cal_events = fetch_events_for_calendar(service, cal_id, time_min_iso, time_max_iso)
        events_total += len(cal_events)

        for evt in cal_events:
            s, e, is_all_day = parse_event_times(evt, tzname)
            overlap = clamp_interval(s, e, week_start, week_end)
            if not overlap:
                continue
            if is_all_day and not count_all_day:
                continue

            cs, ce = overlap
            duration_sec = (ce - cs).total_seconds()

            # explicit event color (1..11) else fall back to that calendar's default (1..24)
            cid = evt.get("colorId") or cal_default_color_id or "default"
            seconds_by_color_id[cid] += duration_sec

    print(f"[INFO] Fetched {events_total} events from {len(calendars)} calendar(s)")

    # Label helper: show hex for event colors and for calendar defaults
    def color_label(cid: str) -> str:
        if cid in event_colors:
            return f"{cid} ({event_colors[cid]})"
        if cid in calendar_colors:
            return f"{cid} (default {calendar_colors[cid]})"
        return f"{cid} (n/a)"

    # Build & print rows
    rows = [
        {"color_id": cid, "color": color_label(cid), "hours": round(secs / 3600.0, 2)}
        for cid, secs in seconds_by_color_id.items()
    ]
    rows.sort(key=lambda r: r["hours"], reverse=True)

    print("\nGoogle Calendar color-hours for last week (Sun–Sat)")
    print(f"From: {week_start:%Y-%m-%d %H:%M}  To: {week_end:%Y-%m-%d %H:%M}")
    print("-" * 60)
    if not rows:
        print("No events found in this window.")
    else:
        total_hours = sum(r["hours"] for r in rows)
        for r in rows:
            pct = (r["hours"] / total_hours * 100.0) if total_hours > 0 else 0.0
            print(f"{r['color']:<24} {r['hours']:>6.2f} h   ({pct:>5.1f}%)")
        print("-" * 60)
        print(f"{'TOTAL':<24} {total_hours:>6.2f} h")

    # Write CSV
    with open(args.csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["color_id", "color", "hours"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n[INFO] Wrote CSV: {args.csv}")


if __name__ == "__main__":
    main()