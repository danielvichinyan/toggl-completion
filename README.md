# Fill Timesheet — Toggl + Azure DevOps

Automatically fills your **Toggl Track** timesheet for a week by combining:

- **Outlook meetings** — read directly from Toggl's calendar integration via the API (no browser/Playwright needed)
- **Azure DevOps tasks** — your **Active** tasks in the current iteration

It places meetings at their real times, fills the rest of each 10:00–18:00 workday with task entries (2–4 tasks/day, 8h total), shows you a table for approval, and only writes to Toggl after you say go.

---

## TL;DR — fill this week in one go

```text
# 1. Open Claude Code in this folder
cd D:\Professional\toggl-completion
claude

# 2. First run only: approve the two MCP servers when prompted, then check they're connected
/mcp

# 3. Fill the timesheet
/fill-timesheet for this week
```

Then review the proposed table and reply **"approve"** (or ask for changes). That's it.

Other invocations:

| Command | Result |
|---|---|
| `/fill-timesheet for this week` | Mon–Fri of the current week |
| `/fill-timesheet for last week` | Mon–Fri of the previous week |
| `/fill-timesheet except Friday` | This week, skipping the named day(s) |
| `/fill-timesheet only Monday, Tuesday` | Only the named day(s) |
| `/fill-timesheet for March 25, 26, 27` | Specific dates |

---

## One-time setup

Everything below is already configured on this machine — this section is for rebuilding from scratch or on a new machine.

### 1. Prerequisites

- **Python 3.13+**, **Node.js** (for `npx`), **Azure CLI** (`az`)
- Python deps:
  ```powershell
  python -m pip install mcp aiohttp python-dotenv
  ```

### 2. Toggl API token

Create `.env` next to `server.py`:

```
TOGGL_API_TOKEN=<your token from Toggl → Profile → API Token>
```

`server.py` loads this `.env` by absolute path, so it works no matter what directory the MCP host launches it from.

### 3. Connect Outlook to Toggl

In Toggl: **Profile → Calendar integrations → Connect** your Outlook account, and select the calendar(s) to show. The skill reads these synced events; you do **not** need "Auto-track calendar events" turned on.

Verify the connection (optional): the `get_calendar_integrations` tool lists connected providers.

### 4. Azure DevOps auth

```powershell
az login
```

The Azure DevOps MCP authenticates as the signed-in `az` user — no PAT required. Org slug is **`spaceflux`** (project **Spaceflux**, team **Spaceflux Team**).

### 5. MCP servers (`.mcp.json`)

Already created in this folder. It registers two project-scoped servers:

```json
{
  "mcpServers": {
    "toggl": {
      "command": "C:\\Users\\DanielVichinyan\\AppData\\Local\\Microsoft\\WindowsApps\\PythonSoftwareFoundation.Python.3.13_qbz5n2kfra8p0\\python.exe",
      "args": ["D:\\Professional\\toggl-completion\\server.py"]
    },
    "azure-devops": {
      "command": "cmd",
      "args": ["/c", "npx", "-y", "@azure-devops/mcp", "spaceflux"]
    }
  }
}
```

Because these are **project-scoped**, Claude Code prompts you to approve them the first time you open this folder. They only load when you run `claude` from here. (Want them everywhere? See "Make it global" below.)

### 6. The skill

Lives at `.claude/skills/fill-timesheet/SKILL.md`. Project-scoped, so it's available as `/fill-timesheet` when you run Claude Code from this folder.

---

## How it works (the 5 steps the skill runs)

1. **Meetings** — `get_calendar_events(start_date, end_date)` (Toggl MCP) returns the week's Outlook meetings, already converted to your timezone, with ready-to-use ISO start/stop strings. All-day events and `Canceled:` events are dropped.
2. **Tasks** — `work_list_team_iterations` finds the iteration covering the week; `wit_get_work_items_for_iteration` + `wit_get_work_items_batch_by_ids` fetch its tasks, filtered to **Active tasks assigned to you**. (In iteration Week 2 it prefers tasks not already logged in Week 1.)
3. **Build** — meetings placed at real times; tasks fill each day to 8h (2–4 entries, 2/3/4h each, one adjusted to make the day total exact).
4. **Approve** — a per-day table is shown; nothing is written until you approve.
5. **Publish** — `create_time_entry` writes each entry to the **Kodin Development** project with the correct timezone offset.

---

## Toggl MCP tools (in `server.py`)

| Tool | Purpose |
|---|---|
| `get_calendar_events` | Read connected-calendar (Outlook) meetings — **replaces browser scraping** |
| `get_calendar_integrations` | List connected calendars / auto-track status |
| `create_time_entry` | Create a finished entry with explicit start/stop |
| `delete_time_entry` | Delete an entry by ID (use to undo) |
| `get_time_entries`, `get_time_summary`, `search_time_entries` | Read/report entries |
| `get_projects`, `get_workspaces`, `get_project_tasks`, `get_all_tasks` | Lookups |
| `start_timer`, `stop_current_timer`, `get_current_timer` | Live timer control |
| `create_project_task` | Create a Toggl project task |

`create_time_entry`, `delete_time_entry`, `get_calendar_events`, and `get_calendar_integrations` are additions on top of the upstream [`toggl-track-mcp`](https://github.com/vontell/toggl-track-mcp).

---

## Gotchas

- **Duplicates:** re-running for the same week creates duplicate entries (no dedupe). To undo, delete the entries by ID via `delete_time_entry`.
- **Calendar sync quota:** reading events (`get_calendar_events`) is a read-only DB call and does **not** consume your rate-limited calendar provider-sync quota (the "X/30 requests"). That quota is only for forcing a fresh pull from Outlook, which this setup never does — Toggl syncs on its own schedule.
- **Timezone:** times follow your Toggl account timezone offset — `+03:00` (EEST, last Sun of Mar → last Sun of Oct) or `+02:00` (EET, otherwise). The calendar tool applies this automatically.
- **Recent meetings missing:** if a meeting you just created in Outlook isn't showing, Toggl hasn't synced it yet — wait for its periodic sync.

---

## Make it global (optional)

To use `/fill-timesheet` from any directory, move the config to your user scope:

- Add the same `mcpServers` block to `~/.claude.json` (user scope), and
- Copy the skill to `~/.claude/skills/fill-timesheet/SKILL.md`.

Then it works regardless of which folder you launch Claude Code from.
