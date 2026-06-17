# toggl-completion — Timesheet Automation

Automatically fills your **Toggl Track** timesheet for a week by combining:

- **Outlook meetings** — read from Toggl's calendar integration (no browser needed)
- **Azure DevOps tasks** — your **Active** tasks in the current sprint iteration

Meetings are placed at their real times. Tasks fill the remaining time in each slot, with a **1-hour minimum per entry** and **10-minute rounding**. Tasks never cross meeting boundaries.

---

## Primary: standalone Python script (`fill_timesheet.py`)

The script runs entirely without Claude or an MCP server. It fetches data, proposes a schedule, waits for approval, submits to Toggl, and optionally closes ADO tasks.

### Quick start

```powershell
# 1. Copy and fill in your credentials
cp .env.example .env
#    edit .env — add TOGGL_API_TOKEN and ADO_PAT

# 2. Install dependencies
pip install requests python-dotenv

# 3. Dry-run (no writes anywhere)
python fill_timesheet.py --dry-run

# 4. Submit for real (asks for confirmation)
python fill_timesheet.py

# 5. Or skip the prompt
python fill_timesheet.py --yes
```

### CLI flags

| Flag | Description |
|---|---|
| `--dry-run` | Print the proposed schedule, make no API calls |
| `--last-week` | Fill Mon–Fri of the previous week |
| `--next-week` | Fill Mon–Fri of next week |
| `--skip MON TUE ...` | Skip these weekdays (e.g. `--skip FRI`) |
| `--only MON TUE ...` | Only fill these weekdays |
| `--yes` / `-y` | Skip the confirmation prompt and submit immediately |
| `--no-close` | Submit Toggl entries but do **not** close ADO tasks |

### How the scheduling works

1. Fetches Toggl calendar events for the target week (Outlook meetings via Toggl integration).
2. Fetches all **Active** ADO tasks assigned to you in the current iteration.
3. Computes free slots between meetings within the 10:00–18:00 work window.
4. Distributes tasks across slots: equal base time per task using largest-remainder rounding, **minimum 60 min** per entry, all durations rounded to **10-minute** multiples.
5. Shows a per-day table and waits for `y/n` confirmation (or skip with `--yes`).
6. Submits entries to Toggl, then closes ADO tasks (unless `--no-close`).

### Credentials

Create `.env` (see `.env.example`):

```
TOGGL_API_TOKEN=...   # Toggl → Profile → API Token
ADO_PAT=...           # Azure DevOps → User Settings → Personal Access Tokens
                      # Required scopes: Work Items (read + write)
```

The `.env` file is git-ignored — never commit it.

**Generating an ADO PAT:** go to `https://dev.azure.com/spaceflux/_usersSettings/tokens`, create a token with **Work Items: Read & Write** scope.

### Hardcoded constants (no lookup overhead)

| Constant | Value |
|---|---|
| Toggl workspace | `6705147` |
| Toggl project (`Kodin Development`) | `185931458` |
| ADO org | `spaceflux` |
| ADO project | `Spaceflux` |
| ADO team | `Spaceflux Team` |
| User email | `daniel.vichinyan@spaceflux.io` |

---

## Secondary: Claude Code MCP skill (`/fill-timesheet`)

An interactive Claude Code skill that does the same thing conversationally. Requires the MCP servers and runs inside `claude`.

```text
cd D:\Professional\toggl-completion
claude
/fill-timesheet for this week
```

Other invocations:

| Command | Result |
|---|---|
| `/fill-timesheet for this week` | Mon–Fri of the current week |
| `/fill-timesheet for last week` | Mon–Fri of the previous week |
| `/fill-timesheet except Friday` | This week, skipping the named day(s) |
| `/fill-timesheet only Monday, Tuesday` | Only the named day(s) |

### MCP setup (one-time)

**1. Install Python deps:**
```powershell
pip install mcp aiohttp python-dotenv
```

**2. Install Node.js** (for `npx`) — [https://nodejs.org](https://nodejs.org)

**3. Toggl MCP server** — `server.py` in this repo wraps the Toggl API.  
It builds on top of the upstream [`toggl-track-mcp`](https://github.com/vontell/toggl-track-mcp) repo (adds `create_time_entry`, `delete_time_entry`, `get_calendar_events`, and `create_time_entries_bulk`). The upstream repo is not required directly — all tools are re-implemented in `server.py`.

**4. Azure DevOps MCP** — installed automatically via `npx`:
```powershell
npx -y @azure-devops/mcp spaceflux
```

**5. Auth:**
- Toggl: `TOGGL_API_TOKEN` in `.env` (same file as the script)
- ADO: `ADO_PAT` in `.env` (same token)

The `.mcp.json` in this folder registers both servers as project-scoped MCP servers. Claude Code prompts you to approve them the first time.

**6. Connect Outlook to Toggl:**  
Toggl → Profile → Calendar integrations → Connect your Outlook account. The skill reads synced events — "Auto-track calendar events" does **not** need to be on.

---

## Toggl MCP tools (in `server.py`)

| Tool | Purpose |
|---|---|
| `get_calendar_events` | Read Outlook meetings from Toggl's calendar integration |
| `get_calendar_integrations` | List connected calendars |
| `create_time_entry` | Create a finished entry with explicit start/stop |
| `create_time_entries_bulk` | Create multiple entries in one call (avoids rate limits) |
| `delete_time_entry` | Delete an entry by ID |
| `get_time_entries`, `get_time_summary`, `search_time_entries` | Read/report entries |
| `get_projects`, `get_workspaces`, `get_project_tasks`, `get_all_tasks` | Lookups |
| `start_timer`, `stop_current_timer`, `get_current_timer` | Live timer control |
| `create_project_task` | Create a Toggl project task |

---

## Gotchas

- **Rate limit:** Toggl free plan allows 30 API requests/hour. The script uses one POST per entry — for a typical week (~15–20 entries) this stays well under the limit.
- **Duplicates:** re-running for the same week creates duplicate entries (no dedupe). Delete via `delete_time_entry` (MCP) or Toggl UI.
- **Calendar sync lag:** if a meeting just created in Outlook isn't showing, Toggl hasn't synced it yet — wait a few minutes.
- **Timezone:** entries use your Toggl account timezone — `+03:00` (EEST, last Sunday of March → last Sunday of October) or `+02:00` (EET, otherwise). The script detects this automatically via the `/me` endpoint.
- **Windows `az` CLI:** the script uses `shell=True` for subprocess calls so `az.cmd` resolves correctly.
