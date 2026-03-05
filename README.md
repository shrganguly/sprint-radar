# Shiproom Dashboard Generator

**Turn your Excel feature tracker into a beautiful, interactive HTML dashboard in under 5 minutes.**

Every shiproom call is the same story -- hunt down the latest deck, manually update bug counts, copy-paste status from five different sources, and hope nothing is stale by the time you present. This tool eliminates all of that. Your feature crew maintains one source of truth: **Excel for features, ADO for bugs**. Run one command, get a fully up-to-date dashboard with live data. No slides to update, no numbers to chase.

Built by the SharePoint team. Now available for every area to use.

---

## What You Get

A single-page HTML dashboard with:

- **Sprint Health** at a glance -- blocked items, rollout status, progress tracking
- **KPI Cards** -- total features, in rollout, shipped, blocked/at-risk counts
- **Bug Burndown** -- live P1/P2 counts, resolved trends (pulled from ADO)
- **IcM Incidents** -- active and resolved incident counts (optional)
- **Epic Trackers** -- collapsible cards grouped by epic with status pills, priority badges, blocker details
- **Resource Utilization** -- per-EM feature load and bug distribution
- **Partner Dependencies** -- risk tracking with impact and mitigation
- **Slipping ETA view** -- items where actual delivery differs from committed

Everything auto-generates from your Excel tracker, Bug ADO board, and IcM queries. No manual HTML editing. Just run the script.

---

## Quick Start (5 minutes)

### Prerequisites

- Windows with Python 3.10+
- Excel installed (script reads via COM automation)
- These Python packages:
  ```
  pip install pywin32 requests msal
  ```
- Optional (for ICM only): `pip install playwright && playwright install`

### Step 1: Get the files

Clone this repo or copy these files to your project folder:

```
your-project/
  generate_html_dashboard.py    # Main script
  _fetch_icm.py                 # ICM scraper (optional)
  config.json                   # Your config (create from template)
  your-feature-tracker.xlsx     # Your Excel tracker
```

### Step 2: Set up your Excel tracker

**Starting fresh?** Copy the template:
```
templates/Feature tracker_template.xlsx
```
It comes pre-configured with all the right tabs, columns, dropdowns, and conditional formatting.

**Already have a tracker?** Make sure it has:
- A **sprint tab** (e.g., "Sprint Mar 2026") with these column headers:
  `Item ID`, `Epic / Feature`, `Work Item`, `Work Type`, `Priority`, `PM Owner`, `Dev Owner`, `Design Owner`, `Sprint Start Status`, `Current Status`, `Target Status`, `Release to Prod Month (Actual)`, `Release to Prod Month (Committed)`, `Blocker`, `Blocker Details`, `Path to Green`, `ETA to Unblock`, `Ask / Decision Needed`, `Link`, `Notes`
- A **Config tab** with a `StatusList` in column A (ordered from earliest to latest stage)
- *(Optional)* `Feature Pitch` column for work item descriptions
- *(Optional)* `Partner Dependencies` tab
- *(Optional)* `EM` tab with EM-to-IC mapping

> The script finds columns by **header name** (case-insensitive, partial match), so column order doesn't matter. Add or rearrange columns freely.

### Step 3: Create your config.json

Copy the template and fill in your values:

```
copy templates\config-template.json config.json
```

**Minimum config** (just Excel, no ADO/ICM):
```json
{
  "dashboard": {
    "title": "Your Team Shiproom View",
    "team_name": "Your Team"
  },
  "paths": {
    "excel_path": "C:\\path\\to\\your-feature-tracker.xlsx"
  }
}
```

That's it. Everything else has sensible defaults.

**With ADO bugs** (add this to enable live bug counts):
```json
{
  "dashboard": {
    "title": "Your Team Shiproom View",
    "team_name": "Your Team"
  },
  "paths": {
    "excel_path": "C:\\path\\to\\your-feature-tracker.xlsx"
  },
  "ado": {
    "enabled": true,
    "organization": "your-ado-org",
    "project": "Your ADO Project",
    "area_path": "Your Project\\Your\\Area\\Path",
    "iteration_path_template": "{project}\\{cy}{half}\\{cy}{quarter}\\{year} - {month_num} {month_short}"
  }
}
```

> ADO uses **MSAL browser login** -- your Microsoft SSO. No PAT tokens to manage. First run opens a browser for login, then it caches the token.

### Step 4: Generate your dashboard

```bash
python generate_html_dashboard.py --config config.json
```

Output: `shiproom-dashboard-MM_DD_YYYY.html` in the same folder as your Excel file.

Open it in any browser. Share it with your team. Done.

---

## Config Reference

The full `config-template.json` has inline `_example` fields explaining every option. Here's the summary:

| Section | Field | Required | What it does |
|---------|-------|----------|-------------|
| **dashboard** | `title` | No | Dashboard heading (default: "Shiproom View") |
| | `team_name` | No | Footer branding |
| | `subtitle` | No | Small text above title |
| **paths** | `excel_path` | **Yes** | Full path to your Excel tracker |
| | `html_output_dir` | No | Where to save HTML (default: same as Excel) |
| **excel** | `sprint_tab_pattern` | No | Tab naming pattern (default: `Sprint {month_short} {year}`) |
| | `dependencies_tab` | No | Dependencies tab name, or `null` to skip |
| | `em_tab` | No | EM tab name, or `null` to skip |
| **ado** | `enabled` | No | `true` to pull live bug counts from ADO |
| | `organization` | If ADO on | Your ADO org (from `dev.azure.com/{org}`) |
| | `project` | If ADO on | Your ADO project name |
| | `area_path` | If ADO on | Your team's area path (double backslashes) |
| | `iteration_path_template` | No | How your iterations are named |
| **icm** | `enabled` | No | `true` to include ICM incident counts |
| | `active_query_url` | If ICM on | ICM saved query URL for active incidents |
| | `resolved_query_url` | If ICM on | ICM saved query URL for resolved incidents |
| **em_layout** | `enabled` | No | `true` to show Resource Utilization per EM |
| | `pairs` | If EM on | Column mapping for each EM in the EM tab |
| **priority_order** | | No | Priority labels and sort order |

> **Status order** and **rolled-out statuses** are read automatically from your Excel Config tab's StatusList column. No need to configure them.

---

## What Gets Enabled When

The dashboard adapts based on your config. You only get sections you've configured:

| You configure... | You get... |
|-----------------|-----------|
| Just `excel_path` | Feature table, epic trackers, sprint health, progress tracking |
| + `ado.enabled` | Bug KPI cards (P1, P2, resolved, opened counts with ADO links) |
| + `icm.enabled` | ICM active/resolved cards in the bug section |
| + `em_layout.enabled` | Resource Utilization tab with per-EM feature + bug breakdown |
| + `dependencies_tab` | Partner Dependencies section in Blockers tab |

Start minimal. Add sections as you need them.

---

## Sprint Tab Pattern

Your sprint tabs can be named however you want. Set `sprint_tab_pattern` using these variables:

| Variable | Example | Description |
|----------|---------|-------------|
| `{month_short}` | Mar | 3-letter month |
| `{month_long}` | March | Full month name |
| `{month_num}` | 03 | Zero-padded month number |
| `{year}` | 2026 | 4-digit year |

Examples:
- `Sprint {month_short} {year}` --> "Sprint Mar 2026" *(default)*
- `{month_long} {year} Sprint` --> "March 2026 Sprint"
- `S{month_num}-{year}` --> "S03-2026"

---

## EM Tab Setup

The EM tab maps ICs to their managers. Layout is side-by-side columns:

```
Col A (EM name)    Col B (Directs)     Col D (EM name)    Col E (Directs)
Dana Martinez      Pat Kumar           Chris Okafor       Morgan Silva
                   Riley Zhang                            Casey Tanaka
                   Drew Patel                             Taylor Kim
```

Config for this layout:
```json
"em_layout": {
  "enabled": true,
  "pairs": [
    { "em_name_row": 2, "em_col": 1, "directs_col": 2 },
    { "em_name_row": 2, "em_col": 4, "directs_col": 5 }
  ]
}
```

Add more entries for more EMs. The script reads until the first blank row in each directs column.

---

## ADO Iteration Path

The `iteration_path_template` tells the script how your ADO iterations are structured. Variables:

| Variable | Example | Description |
|----------|---------|-------------|
| `{project}` | ODSP Product Experiences | Your ADO project name |
| `{cy}` | CY26 | Calendar year prefix |
| `{half}` | H1 | H1 or H2 |
| `{quarter}` | Q1 | Q1-Q4 |
| `{year}` | 2026 | 4-digit year |
| `{month_num}` | 03 | Zero-padded month |
| `{month_short}` | Mar | 3-letter month |

Default: `{project}\{cy}{half}\{cy}{quarter}\{year} - {month_num} {month_short}`
--> `ODSP Product Experiences\CY26H1\CY26Q1\2026 - 03 Mar`

Adjust to match your team's ADO iteration naming convention.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Tab not found" | Check `sprint_tab_pattern` matches your actual tab names |
| ADO shows 0 bugs | Run once to trigger MSAL login. Check `area_path` matches your ADO area exactly |
| MSAL browser doesn't open | Make sure `msal` is installed: `pip install msal` |
| Excel COM error | Close any Excel dialogs or "Protected View" prompts, then retry |
| ICM timeout | ICM requires Edge with your corp login. Run `python _fetch_icm.py --config config.json` manually first to authenticate |
| Wrong column data | Script matches columns by header name. Check your headers match standard names (case-insensitive) |

---

## Claude Code Skill (Optional)

If your team uses [Claude Code](https://claude.com/claude-code), you can install this as a slash command so generating the dashboard is just `/generate-shiproom-view`.

**Install the skill:**
```bash
# Copy the skill folder to your Claude Code skills directory
xcopy /E /I skill "%USERPROFILE%\.claude\skills\generate-shiproom-view"
```

**Use it:**
```
/generate-shiproom-view                     # auto-detects config.json in current directory
/generate-shiproom-view config.json         # explicit path
/generate-shiproom-view C:\path\config.json # absolute path
```

It validates your config, runs the script, and opens the dashboard automatically.

---

## File Structure

```
your-project/
  generate_html_dashboard.py   # Main dashboard generator
  _fetch_icm.py                # ICM incident scraper (optional)
  config.json                  # Your team's config (don't commit — has area paths)
  your-tracker.xlsx            # Your Excel feature tracker (keep in SharePoint)
  shiproom-dashboard-*.html    # Generated output (gitignored)
  .ado_token_cache.bin         # MSAL token cache (gitignored)

templates/
  config-template.json         # Annotated config template with examples
  Feature tracker_template.xlsx # Excel template with formatting + sample data

skill/
  SKILL.md                     # Claude Code skill definition
```

---

## Contributing

Want to customize the dashboard for your team's specific needs? Fork this repo and make it your own -- add new tabs, change the visual style, integrate additional data sources. If you build something that would benefit everyone, send a PR.

---

*Built with Python + win32com + MSAL. No external UI frameworks. One HTML file, zero dependencies to view.*
