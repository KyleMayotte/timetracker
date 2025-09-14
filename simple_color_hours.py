#!/usr/bin/env python3
# Google Calendar color-hours with custom labels (Sun→Sat LAST week)

import os
import argparse
from collections import defaultdict
from datetime import datetime, timedelta, date, time

from dateutil.tz import gettz
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPES = ["https://www.googleapis.com/auth/calendar.readonly"]
TIMEZONE = "America/Chicago"

# Your mapping: Google EVENT color IDs ("1".."11") -> label
CUSTOM_COLOR_NAMES = {
    "1":  "Class",
    "2":  "Side-Hustle",
    "3":  "Personal-Development",
    "4":  "Faith",
    "5":  "Network",
    "6":  "Test",
    "7":  "Money",
    "8":  "Sleep",
    "9":  "Exercise",
    "10": "Project",
    "11": "Meal",
}

# ---------- Auth ----------
def auth_service():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists("credentials.json"):
                raise FileNotFoundError("Missing credentials.json (OAuth Desktop client).")
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            try:
                creds = flow.run_local_server(port=0)
            except Exception:
                print("[INFO] Falling back to console auth.")
                creds = flow.run_console()
        with open("token.json", "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return build("calendar", "v3", credentials=creds)

# ---------- Time window: LAST complete Sunday → Saturday ----------
def get_last_week_window(tzname: str):
    """
    Return [start, end) for the LAST complete Sunday→Saturday week.
    
    If today is Sunday: returns the previous week (7 days ago)
    If today is Monday-Saturday: returns the week that just ended
    
    start = last Sunday 00:00
    end   = last Saturday 23:59:59 + 1 second (i.e., this past Sunday 00:00)
    """
    tz = gettz(tzname)
    now = datetime.now(tz)
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    
    # Find the most recent Sunday (could be today if today is Sunday)
    days_since_sunday = today.weekday()  # Monday=0, Sunday=6
    if days_since_sunday == 6:  # Today is Sunday
        days_since_sunday = 0
    else:
        days_since_sunday = (days_since_sunday + 1) % 7  # Convert to Sunday=0 system
    
    # Get the Sunday that started the week we want to analyze
    if today.weekday() == 6:  # If today is Sunday
        # We want the previous complete week (Sunday 7 days ago to Saturday 1 day ago)
        last_sunday = today - timedelta(days=7)
    else:
        # We want the week that just ended (most recent Sunday to most recent Saturday)
        last_sunday = today - timedelta(days=days_since_sunday)
    
    # Set the time window
    start = last_sunday  # Sunday 00:00:00
    end = start + timedelta(days=7)  # Next Sunday 00:00:00 (exclusive)
    
    return start, end

# ---------- Helpers ----------
def clip(a, b, lo, hi):
    s, e = max(a, lo), min(b, hi)
    return None if s >= e else (s, e)

def parse_event_times(evt, tzname: str):
    tz = gettz(tzname)
    s = evt.get("start", {}); e = evt.get("end", {})
    if "dateTime" in s:
        sd = datetime.fromisoformat(s["dateTime"].replace("Z", "+00:00")).astimezone(tz)
        ed = datetime.fromisoformat(e["dateTime"].replace("Z", "+00:00")).astimezone(tz)
        return sd, ed, False
    # all-day (end date exclusive)
    sd = datetime.combine(date.fromisoformat(s["date"]), time(0,0), tz)
    ed = datetime.combine(date.fromisoformat(e["date"]), time(0,0), tz)
    return sd, ed, True

def hex_to_rgb(hx: str):
    hx = hx.lstrip('#')
    return tuple(int(hx[i:i+2], 16) for i in (0,2,4))

def nearest_event_color_id(hex_color: str, event_hex_map: dict):
    """Map any hex (incl. calendar default color) to the nearest EVENT color id '1'..'11'."""
    try:
        r,g,b = hex_to_rgb(hex_color)
    except Exception:
        return None
    best_id, best_d = None, 10**9
    for cid, ehx in event_hex_map.items():
        if not ehx:
            continue
        er,eg,eb = hex_to_rgb(ehx)
        d = (r-er)**2 + (g-eg)**2 + (b-eb)**2
        if d < best_d:
            best_d, best_id = d, cid
    return best_id

def fetch_all_calendars(service, include_hidden=False):
    cals = []
    page = service.calendarList().list().execute()
    cals.extend(page.get("items", []))
    token = page.get("nextPageToken")
    while token:
        page = service.calendarList().list(pageToken=token).execute()
        cals.extend(page.get("items", []))
        token = page.get("nextPageToken")
    if not include_hidden:
        cals = [c for c in cals if not c.get("hidden")]
    return cals

def fetch_events_for_calendar(service, calendar_id, start_iso, end_iso):
    events = []
    resp = service.events().list(
        calendarId=calendar_id,
        timeMin=start_iso,
        timeMax=end_iso,   # Calendar API: timeMax is EXCLUSIVE (perfect for our window end)
        singleEvents=True,
        orderBy="startTime",
        maxResults=2500
    ).execute()
    events.extend(resp.get("items", []))
    token = resp.get("nextPageToken")
    while token:
        resp = service.events().list(
            calendarId=calendar_id,
            timeMin=start_iso,
            timeMax=end_iso,
            singleEvents=True,
            orderBy="startTime",
            maxResults=2500,
            pageToken=token
        ).execute()
        events.extend(resp.get("items", []))
        token = resp.get("nextPageToken")
    return events

def get_colors_map(service):
    colors = service.colors().get().execute()
    event_colors = {k: v.get("background") for k, v in colors.get("event", {}).items()}       # "1".."11" -> hex
    calendar_colors = {k: v.get("background") for k, v in colors.get("calendar", {}).items()} # "1".."24" -> hex
    return event_colors, calendar_colors

def format_date_range(start, end):
    """Format the date range nicely for display."""
    end_display = end - timedelta(seconds=1)  # Show the last second of Saturday
    return f"{start.strftime('%A, %B %d')} to {end_display.strftime('%A, %B %d, %Y')}"

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="Print hours per custom label for the LAST complete Sunday→Saturday week.")
    ap.add_argument("--tz", default=TIMEZONE, help="IANA timezone, e.g. America/Chicago")
    ap.add_argument("--count-all-day", action="store_true", help="Count all-day events as 24h/day.")
    ap.add_argument("--all-calendars", action="store_true", help="Sum across all calendars you can see.")
    ap.add_argument("--csv", help="Also write CSV to this file.")
    args = ap.parse_args()

    service = auth_service()

    # Get last complete Sunday→Saturday week
    start, end = get_last_week_window(args.tz)
    date_range = format_date_range(start, end)
    print(f"[INFO] Analyzing most recently completed week: {date_range}")
    print(f"[INFO] Time window: {start} → {end} ({args.tz})")

    # Resolve Google color tables
    event_colors, calendar_colors = get_colors_map(service)

    # Which calendars to scan
    if args.all_calendars:
        calendars = fetch_all_calendars(service)
        print(f"[INFO] Scanning {len(calendars)} calendars")
    else:
        calendars = [service.calendarList().get(calendarId="primary").execute()]
        print(f"[INFO] Scanning primary calendar only")

    seconds_by_label = defaultdict(float)
    total_events = 0

    for cal in calendars:
        cal_id = cal["id"]
        cal_name = cal.get("summary", cal_id)
        cal_default_id = cal.get("colorId")           # e.g., "14"
        cal_default_hex = calendar_colors.get(cal_default_id) if cal_default_id else None

        cal_events = fetch_events_for_calendar(service, cal_id, start.isoformat(), end.isoformat())
        if cal_events:
            print(f"[INFO] Found {len(cal_events)} events in '{cal_name}'")
        total_events += len(cal_events)

        for evt in cal_events:
            s, e, is_all_day = parse_event_times(evt, args.tz)
            clipped = clip(s, e, start, end)
            if not clipped:
                continue
            if is_all_day and not args.count_all_day:
                continue

            cs, ce = clipped
            dur = (ce - cs).total_seconds()

            # Decide label
            cid = evt.get("colorId")  # explicit event color id "1".."11"
            if cid and cid in CUSTOM_COLOR_NAMES:
                label = CUSTOM_COLOR_NAMES[cid]
            else:
                # Uncolored event → map the calendar's default hex to the nearest event color id
                if cal_default_hex:
                    nearest_id = nearest_event_color_id(cal_default_hex, event_colors)
                    label = CUSTOM_COLOR_NAMES.get(nearest_id, "Miscellaneous")
                else:
                    label = "Miscellaneous"

            seconds_by_label[label] += dur

    print(f"\n[INFO] Total events processed: {total_events}")

    if not seconds_by_label:
        print("No events found in this time window.")
        if args.csv:
            import csv
            with open(args.csv, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["label", "hours"])
        return

    rows = [(label, round(secs/3600.0, 2)) for label, secs in seconds_by_label.items()]
    rows.sort(key=lambda x: x[1], reverse=True)

    print(f"\n=== Hours by Label ({date_range}) ===")
    total_hours = sum(hrs for _, hrs in rows)
    for name, hrs in rows:
        percentage = (hrs / total_hours * 100) if total_hours > 0 else 0
        print(f"{name:20}: {hrs:6.2f} hours ({percentage:5.1f}%)")
    
    print(f"{'='*50}")
    print(f"{'Total':20}: {total_hours:6.2f} hours")
    print("=" * 50)

    if args.csv:
        import csv
        with open(args.csv, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["label","hours","percentage"])
            for name, hrs in rows:
                percentage = (hrs / total_hours * 100) if total_hours > 0 else 0
                w.writerow([name, hrs, f"{percentage:.1f}%"])
        print(f"\n[INFO] CSV report saved to: {args.csv}")

if __name__ == "__main__":
    main()