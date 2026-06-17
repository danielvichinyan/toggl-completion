#!/usr/bin/env python3
"""
fill_timesheet.py — Fill Toggl timesheet from Outlook calendar + Azure DevOps active tasks.

Usage:
  python fill_timesheet.py               # fill this week (Mon–Fri)
  python fill_timesheet.py --last-week   # fill last week
  python fill_timesheet.py --skip mon,fri    # skip specific days
  python fill_timesheet.py --only wed,thu    # only fill specific days
  python fill_timesheet.py --dry-run     # show proposal, don't submit
"""

import argparse
import base64
import json
import os
import random
import re
import subprocess
import sys
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from datetime import date, datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

# ── Constants ──────────────────────────────────────────────────────────────────

TOGGL_API_BASE   = "https://api.track.toggl.com/api/v9"
TOGGL_WORKSPACE  = 6705147
TOGGL_PROJECT_ID = 185931458
TOGGL_PROJECT    = "Kodin Development"

ADO_ORG     = "spaceflux"
ADO_PROJECT = "Spaceflux"
ADO_TEAM    = "Spaceflux Team"
ADO_BASE    = f"https://dev.azure.com/{ADO_ORG}"
ADO_RESOURCE_ID = "499b84ac-1321-427f-aa17-267ca6975798"

USER_EMAIL  = "daniel.vichinyan@spaceflux.io"
WORK_START  = time(10, 0)
WORK_END    = time(18, 0)
DAY_NAMES   = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]

# Resolved at startup from Toggl /me (avoids hardcoding DST offset)
TZ: ZoneInfo = None  # type: ignore

# ── Auth helpers ───────────────────────────────────────────────────────────────

def toggl_headers() -> dict:
    token = os.getenv("TOGGL_API_TOKEN")
    if not token:
        sys.exit("TOGGL_API_TOKEN not set in .env")
    encoded = base64.b64encode(f"{token}:api_token".encode()).decode()
    return {"Authorization": f"Basic {encoded}", "Content-Type": "application/json"}


def resolve_tz() -> ZoneInfo:
    """Fetch the Toggl account timezone once and cache it."""
    global TZ
    if TZ is not None:
        return TZ
    r = requests.get(f"{TOGGL_API_BASE}/me", headers=toggl_headers())
    if r.status_code == 402:
        sys.exit(f"Toggl rate limit hit: {r.text.strip()}")
    r.raise_for_status()
    tz_name = r.json().get("timezone") or "UTC"
    try:
        TZ = ZoneInfo(tz_name)
    except Exception:
        TZ = ZoneInfo("UTC")
    return TZ


_ado_token_cache: str = ""


def ado_headers() -> dict:
    """
    Build ADO auth headers. Uses ADO_PAT from .env if set (Basic auth — most reliable).
    Falls back to az account get-access-token Bearer token.
    """
    global _ado_token_cache
    if _ado_token_cache:
        return {"Authorization": f"Bearer {_ado_token_cache}", "Content-Type": "application/json"}

    pat = os.getenv("ADO_PAT", "").strip()
    if pat:
        encoded = base64.b64encode(f":{pat}".encode()).decode()
        return {"Authorization": f"Basic {encoded}", "Content-Type": "application/json"}

    # Fallback: Azure CLI Bearer token (may 401 depending on org AAD setup)
    result = subprocess.run(
        f"az account get-access-token --resource {ADO_RESOURCE_ID} --query accessToken -o tsv",
        capture_output=True, text=True, shell=True
    )
    if result.returncode != 0:
        sys.exit(f"Failed to get ADO token — set ADO_PAT in .env or run 'az login': "
                 f"{result.stderr.strip() or result.stdout.strip()}")
    _ado_token_cache = result.stdout.strip()
    return {"Authorization": f"Bearer {_ado_token_cache}", "Content-Type": "application/json"}


# ── Date helpers ───────────────────────────────────────────────────────────────

def week_monday(ref: date = None, offset_weeks: int = 0) -> date:
    d = ref or date.today()
    return d - timedelta(days=d.weekday()) + timedelta(weeks=offset_weeks)


def parse_day_names(s: str) -> set[int]:
    """'mon,wed,fri' → {0, 2, 4}"""
    result = set()
    for part in s.lower().split(","):
        part = part.strip()
        if part in DAY_NAMES:
            result.add(DAY_NAMES.index(part))
        else:
            sys.exit(f"Unknown day name: {part!r}. Use mon/tue/wed/thu/fri.")
    return result


def local_dt(dt_str: str) -> datetime:
    """Parse ISO string → datetime in resolved TZ."""
    return datetime.fromisoformat(dt_str.replace("Z", "+00:00")).astimezone(resolve_tz())


def format_hm(dt: datetime) -> str:
    return dt.strftime("%H:%M")


def day_heading(d: date) -> str:
    """'Monday, June 16' without platform-specific %-d."""
    return d.strftime("%A, %B ") + str(d.day)


def duration_str(minutes: int) -> str:
    h, m = divmod(int(minutes), 60)
    return f"{h}h{m:02d}m" if m else f"{h}h"


# ── Toggl API ──────────────────────────────────────────────────────────────────

def toggl_get(path: str, params: dict = None) -> any:
    r = requests.get(f"{TOGGL_API_BASE}{path}", headers=toggl_headers(), params=params)
    if r.status_code == 402:
        sys.exit(f"Toggl rate limit hit: {r.text.strip()}")
    r.raise_for_status()
    return r.json()


def get_calendar_events(start: date, end: date) -> dict[date, list[dict]]:
    """Return non-canceled, non-all-day events grouped by local date."""
    all_events = []
    page_token = None
    while True:
        params = {"start_date": str(start), "end_date": str(end), "limit": 250}
        if page_token:
            params["page_token"] = page_token
        r = requests.get(
            f"{TOGGL_API_BASE}/integrations/calendar/events",
            headers=toggl_headers(), params=params
        )
        if r.status_code == 402:
            sys.exit(f"Toggl rate limit hit: {r.text.strip()}")
        r.raise_for_status()
        data = r.json()
        all_events.extend(data.get("events") or [])
        page_token = data.get("next_page_token")
        if not page_token:
            break

    by_date: dict[date, list[dict]] = {}
    for ev in all_events:
        title = ev.get("title", "")
        if title.startswith("Canceled:") or ev.get("all_day"):
            continue
        start_iso = ev.get("start_time", "")
        end_iso   = ev.get("end_time", "")
        if not start_iso:
            continue
        s = local_dt(start_iso)
        e = local_dt(end_iso) if end_iso else s
        d = s.date()
        by_date.setdefault(d, []).append({
            "title": title, "start": s, "stop": e,
            "start_iso": s.isoformat(), "stop_iso": e.isoformat(),
        })
    return by_date


def get_existing_entries(start: date, end: date) -> list[dict]:
    params = {"start_date": str(start), "end_date": str(end)}
    return toggl_get("/me/time_entries", params) or []


def create_entries_bulk(entries: list[dict]) -> list[str]:
    """
    Create all time entries in one HTTP session (workspace + project fetched once).
    Returns list of result lines.
    """
    hdrs = toggl_headers()
    results = []
    for e in entries:
        from datetime import datetime as dt
        s = dt.fromisoformat(e["start_iso"])
        stop = dt.fromisoformat(e["stop_iso"])
        duration = int((stop - s).total_seconds())
        payload = {
            "created_with": "fill_timesheet.py",
            "description":  e["description"],
            "workspace_id": TOGGL_WORKSPACE,
            "project_id":   TOGGL_PROJECT_ID,
            "duration":     duration,
            "start":        e["start_iso"],
            "stop":         e["stop_iso"],
            "tags":         [],
            "billable":     False,
        }
        r = requests.post(
            f"{TOGGL_API_BASE}/workspaces/{TOGGL_WORKSPACE}/time_entries",
            headers=hdrs, json=payload
        )
        if r.status_code == 402:
            sys.exit(f"Toggl rate limit hit after {len(results)} entries: {r.text.strip()}")
        elif r.status_code in (200, 201):
            results.append(f"  OK  {e['description'][:60]}")
        else:
            results.append(f"  ERR {e['description'][:60]} -> {r.status_code}: {r.text[:80]}")
    return results


# ── ADO API ────────────────────────────────────────────────────────────────────

def ado_get(path: str, params: dict = None) -> any:
    url = f"{ADO_BASE}{path}"
    r = requests.get(url, headers=ado_headers(), params={**(params or {}), "api-version": "7.1"})
    r.raise_for_status()
    return r.json()


def ado_post(path: str, body: dict) -> any:
    url = f"{ADO_BASE}{path}"
    r = requests.post(url, headers=ado_headers(), params={"api-version": "7.1"}, json=body)
    r.raise_for_status()
    return r.json()


def get_iteration(target_monday: date) -> dict:
    data = ado_get(f"/{ADO_PROJECT}/{ADO_TEAM}/_apis/work/teamsettings/iterations")
    for it in (data.get("value") or []):
        attrs = it.get("attributes") or {}
        if not attrs.get("startDate") or not attrs.get("finishDate"):
            continue
        start  = date.fromisoformat(attrs["startDate"][:10])
        finish = date.fromisoformat(attrs["finishDate"][:10])
        if start <= target_monday <= finish:
            return it
    sys.exit(f"Could not find an iteration containing {target_monday}.")


def get_active_tasks(iteration_id: str) -> list[dict]:
    """Return Active tasks assigned to USER_EMAIL in this iteration."""
    wi_data = ado_get(f"/{ADO_PROJECT}/{ADO_TEAM}/_apis/work/teamsettings/iterations/{iteration_id}/workitems")
    relations = wi_data.get("workItemRelations") or []

    # Child task IDs only
    child_ids = [
        r["target"]["id"] for r in relations
        if r.get("rel") == "System.LinkTypes.Hierarchy-Forward"
    ]
    if not child_ids:
        return []

    # Batch fetch in chunks of 200
    fields = ["System.Id", "System.Title", "System.WorkItemType",
              "System.AssignedTo", "System.State"]
    tasks = []
    for i in range(0, len(child_ids), 200):
        chunk = child_ids[i:i+200]
        body = {"ids": chunk, "fields": fields}
        data = ado_post("/_apis/wit/workitemsbatch", body)
        for wi in data.get("value") or []:
            f = wi.get("fields", {})
            assigned = f.get("System.AssignedTo") or {}
            # AssignedTo is a user object; uniqueName holds the email
            assigned_email = assigned.get("uniqueName", "") if isinstance(assigned, dict) else str(assigned)
            if USER_EMAIL not in assigned_email:
                continue
            if f.get("System.State") != "Active":
                continue
            tasks.append({"id": f["System.Id"], "title": f["System.Title"]})
    return tasks


def close_tasks(task_ids: list[int]) -> list[str]:
    """Set System.State = Closed for each task ID. Returns result lines."""
    hdrs = ado_headers()
    hdrs["Content-Type"] = "application/json-patch+json"
    results = []
    patch = [{"op": "replace", "path": "/fields/System.State", "value": "Closed"}]
    for tid in task_ids:
        r = requests.patch(
            f"{ADO_BASE}/_apis/wit/workItems/{tid}",
            headers=hdrs, params={"api-version": "7.1"}, json=patch
        )
        if r.status_code in (200, 201):
            results.append(f"  OK  #{tid} -> Closed")
        else:
            results.append(f"  ERR #{tid} -> {r.status_code}: {r.text[:80]}")
    return results


# ── Scheduling ─────────────────────────────────────────────────────────────────

def free_slots(d: date, meetings: list[dict]) -> list[tuple[datetime, datetime]]:
    """Return free (start, end) pairs within the workday, skipping meetings."""
    tz = resolve_tz()
    day_start = datetime.combine(d, WORK_START, tzinfo=tz)
    day_end   = datetime.combine(d, WORK_END,   tzinfo=tz)
    sorted_m  = sorted(meetings, key=lambda m: m["start"])
    slots = []
    cursor = day_start
    for m in sorted_m:
        if m["start"] > cursor:
            slots.append((cursor, m["start"]))
        cursor = max(cursor, m["stop"])
    if cursor < day_end:
        slots.append((cursor, day_end))
    # Drop slots < 15 min
    return [(s, e) for s, e in slots if (e - s).total_seconds() >= 900]


MIN_ENTRY_MIN = 60  # hard minimum: no Toggl entry shorter than 1 hour


def build_schedule(
    days: list[date],
    events_by_day: dict[date, list[dict]],
    tasks: list[dict],
    used_ids: set[int],
) -> list[dict]:
    """
    Distribute ALL tasks across the week's free slots.
    Tasks are assigned to slots (not days) so no task ever crosses a meeting
    boundary — eliminating sub-1h fragments entirely.
    Each slot's time is divided equally in 10-min increments among its tasks.
    Hard minimum: every entry >= MIN_ENTRY_MIN (60 min).
    Unused-in-Week-1 tasks are ordered first.
    """
    if not tasks:
        sys.exit("No Active tasks assigned to you in this iteration.")

    ordered = ([t for t in tasks if t["id"] not in used_ids]
               + [t for t in tasks if t["id"] in used_ids])
    N = len(ordered)

    # Precompute meetings and slots per day
    day_meetings = {d: sorted(events_by_day.get(d, []), key=lambda m: m["start"]) for d in days}
    day_slots    = {d: free_slots(d, day_meetings[d]) for d in days}

    # Flatten all slots; only keep those long enough for at least 1 task (>= 1h)
    all_slots: list[tuple[datetime, datetime, date, int]] = []
    for d in days:
        for s, e in day_slots[d]:
            dur_10 = (int((e - s).total_seconds() / 60) // 10) * 10
            if dur_10 >= MIN_ENTRY_MIN:
                all_slots.append((s, e, d, dur_10))

    if not all_slots:
        sys.exit("No free slots long enough to schedule tasks (need >= 1h per slot).")

    total_10 = sum(sl[3] for sl in all_slots)

    # Base per-task allocation (10-min multiple, at least MIN_ENTRY_MIN)
    per_task = max(MIN_ENTRY_MIN, (total_10 // N // 10) * 10)

    # Assign task counts to slots proportionally, capped so each task gets >= 1h
    raw     = [sl[3] / per_task for sl in all_slots]
    floored = [min(int(r), sl[3] // MIN_ENTRY_MIN) for r, sl in zip(raw, all_slots)]
    assigned = sum(floored)

    # Distribute remaining tasks to slots with most spare capacity (largest remainder)
    fracs = [(raw[i] - floored[i], i) for i in range(len(all_slots))
             if floored[i] < all_slots[i][3] // MIN_ENTRY_MIN]
    fracs.sort(reverse=True)
    for _, i in fracs:
        if assigned >= N:
            break
        floored[i] += 1
        assigned   += 1

    # If still short (edge case: per_task too large), top up last eligible slots
    i = 0
    while assigned < N and i < len(all_slots):
        cap = all_slots[i][3] // MIN_ENTRY_MIN
        add = min(N - assigned, cap - floored[i])
        floored[i]  += add
        assigned    += add
        i += 1

    # Assign tasks to slots in order
    pool = list(ordered)
    slot_tasks: list[list[dict]] = []
    for i, count in enumerate(floored):
        slot_tasks.append(pool[:count])
        pool = pool[count:]
    if pool:
        slot_tasks[-1].extend(pool)

    # Build per-slot entries (all within one slot — no cross-meeting splits)
    day_entries: dict[date, list[dict]] = {d: [] for d in days}
    for (slot_start, slot_end, d, slot_dur_10), tasks_here in zip(all_slots, slot_tasks):
        if not tasks_here:
            continue
        n    = len(tasks_here)
        base = (slot_dur_10 // n // 10) * 10
        base = max(MIN_ENTRY_MIN, base)
        extra = min(n, (slot_dur_10 - n * base) // 10)
        task_mins = [base + (10 if j < extra else 0) for j in range(n)]

        cursor = slot_start
        for task, tmins in zip(tasks_here, task_mins):
            entry_stop = cursor + timedelta(minutes=tmins)
            day_entries[d].append({
                "is_task":     True,
                "task_id":     task["id"],
                "description": f"#{task['id']} - {task['title']}",
                "start_iso":   cursor.isoformat(),
                "stop_iso":    entry_stop.isoformat(),
            })
            cursor = entry_stop

    # Merge with meetings and sort
    schedule = []
    for d in days:
        meeting_entries = [
            {"is_task": False, "description": m["title"],
             "start_iso": m["start_iso"], "stop_iso": m["stop_iso"]}
            for m in day_meetings[d]
        ]
        all_items = sorted(meeting_entries + day_entries[d], key=lambda x: x["start_iso"])
        schedule.append({"date": d, "entries": all_items})

    return schedule


# ── Display ────────────────────────────────────────────────────────────────────

def display_schedule(schedule: list[dict]) -> None:
    W = 72  # inner content width (excluding leading "  ")
    DESC_MAX = W - 26  # space left after time + dur columns

    grand_total = 0
    for day in schedule:
        d = day["date"]
        print(f"\n  {day_heading(d)}")
        print("  " + "─" * W)
        print(f"  {'TIME':<13}  {'DUR':>5}  DESCRIPTION")
        print("  " + "─" * W)
        day_min = 0
        for e in day["entries"]:
            s    = datetime.fromisoformat(e["start_iso"])
            stop = datetime.fromisoformat(e["stop_iso"])
            mins = int((stop - s).total_seconds() / 60)
            time_str = f"{format_hm(s)}–{format_hm(stop)}"
            dur  = duration_str(mins)
            marker = "◆" if not e["is_task"] else " "
            desc = e["description"]
            if len(desc) > DESC_MAX:
                desc = desc[:DESC_MAX - 3] + "..."
            print(f"  {time_str:<13}  {dur:>5}  {marker} {desc}")
            day_min += mins
        print("  " + "─" * W)
        print(f"  {'Day total:':<30} {duration_str(day_min):>6}")
        grand_total += day_min

    print()
    print("  " + "═" * W)
    print(f"  {'Week total:':<30} {duration_str(grand_total):>6}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--last-week", action="store_true", help="Fill last week instead of this week")
    parser.add_argument("--next-week", action="store_true", help="Fill next week instead of this week")
    parser.add_argument("--skip",  metavar="DAYS", help="Comma-separated days to skip, e.g. mon,fri")
    parser.add_argument("--only",  metavar="DAYS", help="Only fill these days, e.g. wed,thu")
    parser.add_argument("--dry-run", action="store_true", help="Show proposal without submitting to Toggl/ADO")
    parser.add_argument("--yes", "-y", action="store_true", help="Skip approval prompt and submit immediately")
    parser.add_argument("--no-close", action="store_true", help="Skip closing ADO tasks after submitting to Toggl")
    args = parser.parse_args()

    # ── Determine target days ──
    offset = -1 if args.last_week else (1 if args.next_week else 0)
    monday = week_monday(offset_weeks=offset)
    weekdays = [monday + timedelta(days=i) for i in range(5)]

    if args.only:
        keep = parse_day_names(args.only)
        weekdays = [d for d in weekdays if d.weekday() in keep]
    if args.skip:
        skip = parse_day_names(args.skip)
        weekdays = [d for d in weekdays if d.weekday() not in skip]

    if not weekdays:
        sys.exit("No days selected.")

    start_date = weekdays[0]
    end_date   = weekdays[-1]
    week_label = f"{start_date.strftime('%b')} {start_date.day}–{end_date.strftime('%b')} {end_date.day}, {end_date.year}"
    print(f"\nFilling timesheet for {week_label} ({len(weekdays)} day(s))…")

    # ── Calendar events ──
    print("  Fetching calendar events from Toggl…")
    # end_date +1 to ensure Friday events are included
    events_by_day = get_calendar_events(start_date, end_date + timedelta(days=1))

    # ── Iteration + tasks ──
    print("  Fetching ADO iteration…")
    iteration = get_iteration(monday)
    it_attrs  = iteration.get("attributes", {})
    it_start  = date.fromisoformat(it_attrs["startDate"][:10])
    it_end    = date.fromisoformat(it_attrs["finishDate"][:10])
    it_name   = iteration.get("name", "")
    it_id     = iteration["id"]

    week1_monday = it_start
    week2_monday = it_start + timedelta(days=7)
    is_week2 = monday >= week2_monday
    print(f"  Iteration: {it_name} ({it_start}–{it_end}), Week {'2' if is_week2 else '1'}")

    print("  Fetching Active tasks assigned to you…")
    tasks = get_active_tasks(it_id)
    if not tasks:
        # Target iteration has no tasks — fall back to today's iteration
        today_iter = get_iteration(week_monday())
        today_id   = today_iter["id"]
        if today_id != it_id:
            today_name = today_iter.get("name", "current iteration")
            print(f"  No tasks in {iteration.get('name')} — using tasks from {today_name}…")
            tasks = get_active_tasks(today_id)
    if not tasks:
        sys.exit("No Active tasks assigned to you in this or the current iteration.")
    print(f"  Found {len(tasks)} active tasks.")

    # ── Week-2: find already-used task IDs from week 1 Toggl entries ──
    used_ids: set[int] = set()
    if is_week2:
        print("  Week 2: checking which tasks were used in Week 1…")
        wk1_entries = get_existing_entries(week1_monday, week2_monday - timedelta(days=1))
        for entry in wk1_entries:
            m = re.search(r"#(\d+)", entry.get("description", ""))
            if m:
                used_ids.add(int(m.group(1)))
        print(f"  Tasks already used in Week 1: {sorted(used_ids)}")

    # ── Build schedule ──
    print("  Building schedule…")
    random.seed()  # true randomness each run
    schedule = build_schedule(weekdays, events_by_day, tasks, used_ids)

    # ── Display proposal ──
    W = 72
    print(f"\n  {'═' * W}")
    print(f"  PROPOSED TIMESHEET — {week_label}")
    print(f"  Project: {TOGGL_PROJECT}    ◆ = meeting")
    print(f"  {'═' * W}")
    display_schedule(schedule)

    if args.dry_run:
        print("\n  [dry-run] Nothing submitted.")
        return

    # ── Approval ──
    if not args.yes:
        answer = input("\nApprove and submit? [y/N] ").strip().lower()
        if answer != "y":
            print("Aborted.")
            return

    # ── Publish to Toggl ──
    print("\nCreating Toggl entries…")
    all_entries = [e for day in schedule for e in day["entries"]]
    results = create_entries_bulk(all_entries)
    failed = [r for r in results if "OK" not in r]
    for line in failed:
        print(f"  FAILED: {line}")

    # ── Close ADO tasks ──
    task_ids = list({e["task_id"] for day in schedule for e in day["entries"] if e["is_task"]})
    close_results: list[str] = []
    if task_ids and not args.no_close:
        print(f"Closing {len(task_ids)} ADO tasks…")
        close_results = close_tasks(task_ids)
        for line in close_results:
            if "OK" not in line:
                print(f"  FAILED: {line}")
    elif args.no_close:
        print(f"Skipping ADO close (--no-close). {len(task_ids)} tasks left Active.")

    # ── Final summary ──
    W = 72
    toggl_ok = sum(1 for r in results if "OK" in r)
    ado_ok   = sum(1 for r in close_results if "OK" in r)
    print(f"\n  {'═' * W}")
    status = f"{toggl_ok}/{len(results)} entries logged"
    if task_ids:
        status += f"  ·  {ado_ok}/{len(task_ids)} tasks closed" if not args.no_close else f"  ·  {len(task_ids)} tasks left Active"
    print(f"  SUBMITTED — {week_label}    {status}")
    print(f"  {'═' * W}")

    task_entries = [e for day in schedule for e in day["entries"] if e["is_task"]]
    print(f"\n  {'#':>6}  {'TIME':<13}  {'DUR':>5}  TASK")
    print("  " + "─" * W)
    for i, e in enumerate(task_entries, 1):
        s    = datetime.fromisoformat(e["start_iso"])
        stop = datetime.fromisoformat(e["stop_iso"])
        mins = int((stop - s).total_seconds() / 60)
        time_str = f"{format_hm(s)}–{format_hm(stop)}"
        desc = e["description"]
        max_desc = W - 30
        if len(desc) > max_desc:
            desc = desc[:max_desc - 3] + "..."
        print(f"  {i:>6}  {time_str:<13}  {duration_str(mins):>5}  {desc}")
    print("  " + "─" * W)
    print(f"  {len(task_entries):>6} tasks  {'':13}  {duration_str(sum(int((datetime.fromisoformat(e['stop_iso']) - datetime.fromisoformat(e['start_iso'])).total_seconds() / 60) for e in task_entries)):>5}  total task time")


if __name__ == "__main__":
    main()
