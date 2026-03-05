"""
Generate Shiproom Dashboard as HTML.
Config-driven — reads settings from a JSON config file.
Supports: Excel tracker, ADO bugs, ICM incidents, EM resource utilization.
"""
import sys
sys.stdout.reconfigure(encoding='utf-8')
import argparse
import win32com.client
import os
import json
import requests
from html import escape
from urllib.parse import quote
from datetime import datetime
from collections import Counter, defaultdict
import pythoncom
pythoncom.CoInitialize()

# ============================================================
# CONFIG — loaded from JSON file
# ============================================================
parser = argparse.ArgumentParser(description="Generate Shiproom Dashboard HTML from config.")
parser.add_argument("--config", required=True, help="Path to config.json")
args = parser.parse_args()

def load_config(path):
    """Load config JSON, apply defaults, validate required fields."""
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # Strip keys starting with '_' (comment/example fields) recursively
    def strip_comments(obj):
        if isinstance(obj, dict):
            return {k: strip_comments(v) for k, v in obj.items() if not k.startswith("_")}
        if isinstance(obj, list):
            return [strip_comments(i) for i in obj]
        return obj
    cfg = strip_comments(cfg)

    # Apply defaults for optional fields
    dash = cfg.setdefault("dashboard", {})
    dash.setdefault("title", "Shiproom View")
    dash.setdefault("subtitle", "Sprint Review Dashboard")
    dash.setdefault("team_name", "")
    dash.setdefault("source_label", "")

    paths = cfg.setdefault("paths", {})
    paths.setdefault("html_output_dir", None)
    paths.setdefault("bug_data_fallback_path", None)

    excel = cfg.setdefault("excel", {})
    excel.setdefault("sprint_tab_pattern", "Sprint {month_short} {year}")
    excel.setdefault("dependencies_tab", None)
    excel.setdefault("em_tab", None)

    ado = cfg.setdefault("ado", {})
    ado.setdefault("enabled", False)
    ado.setdefault("organization", "")
    ado.setdefault("project", "")
    ado.setdefault("area_path", "")
    ado.setdefault("iteration_path_template", "{project}\\{cy}{half}\\{cy}{quarter}\\{year} - {month_num} {month_short}")

    icm = cfg.setdefault("icm", {})
    icm.setdefault("enabled", False)
    icm.setdefault("active_query_url", "")
    icm.setdefault("resolved_query_url", "")
    icm.setdefault("data_file", "icm_data.json")

    em = cfg.setdefault("em_layout", {})
    em.setdefault("enabled", False)
    em.setdefault("pairs", [])

    cfg.setdefault("priority_order", {"P0": 0, "P1": 1, "P2": 2, "P3": 3})

    # Validate required fields
    if not paths.get("excel_path"):
        print("ERROR: paths.excel_path is required in config.")
        sys.exit(1)
    if ado["enabled"] and not ado.get("area_path"):
        print("ERROR: ado.area_path is required when ado.enabled is true.")
        sys.exit(1)

    return cfg

cfg = load_config(args.config)

# Extract config into module-level variables
EXCEL_PATH = os.path.abspath(cfg["paths"]["excel_path"])
HTML_OUTPUT_DIR = os.path.abspath(cfg["paths"]["html_output_dir"]) if cfg["paths"]["html_output_dir"] else os.path.dirname(EXCEL_PATH)
HTML_OUTPUT_PATH = os.path.join(HTML_OUTPUT_DIR, f"shiproom-dashboard-{datetime.now().strftime('%m_%d_%Y')}.html")
BUG_DATA_PATH = cfg["paths"]["bug_data_fallback_path"]

DASHBOARD_TITLE = cfg["dashboard"]["title"]
DASHBOARD_SUBTITLE = cfg["dashboard"]["subtitle"]
TEAM_NAME = cfg["dashboard"]["team_name"]
SOURCE_LABEL = cfg["dashboard"]["source_label"] or os.path.splitext(os.path.basename(EXCEL_PATH))[0]

ADO_ENABLED = cfg["ado"]["enabled"]
ADO_ORG = cfg["ado"]["organization"]
ADO_PROJECT = cfg["ado"]["project"]
ADO_AREA_PATH = cfg["ado"]["area_path"]
ADO_ITERATION_TEMPLATE = cfg["ado"]["iteration_path_template"]
ADO_BASE = f"https://dev.azure.com/{quote(ADO_ORG)}/{quote(ADO_PROJECT)}/_workitems" if ADO_ENABLED else ""
ADO_API_BASE = f"https://dev.azure.com/{quote(ADO_ORG)}/{quote(ADO_PROJECT)}/_apis" if ADO_ENABLED else ""

def get_ado_auth_headers():
    """Get ADO auth headers via MSAL interactive browser login."""
    import msal
    ADO_RESOURCE = "499b84ac-1321-427f-aa17-267ca6975798"
    CLIENT_ID = "872cd9fa-d31f-45e0-9eab-6e460a02d1f1"
    AUTHORITY = "https://login.microsoftonline.com/organizations"
    SCOPES = [f"{ADO_RESOURCE}/.default"]

    # Cache tokens so you only auth once per session
    _cache_path = os.path.join(os.path.dirname(os.path.abspath(args.config)), ".ado_token_cache.bin")
    cache = msal.SerializableTokenCache()
    if os.path.exists(_cache_path):
        cache.deserialize(open(_cache_path, "r").read())

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)

    # Try silent auth first (cached token)
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        print("  ADO auth: opening browser for Microsoft login...")
        result = app.acquire_token_interactive(scopes=SCOPES)

    # Save cache
    if cache.has_state_changed:
        with open(_cache_path, "w") as f:
            f.write(cache.serialize())

    if "access_token" in result:
        token = result["access_token"]
        print(f"  ADO auth: MSAL token acquired (expires: {result.get('expires_in', '?')}s)")
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
    else:
        print(f"  ADO auth ERROR: {result.get('error')}: {result.get('error_description')}")
        raise Exception(f"MSAL auth failed: {result.get('error_description')}")

ICM_ENABLED = cfg["icm"]["enabled"]
ICM_DATA_PATH = os.path.join(HTML_OUTPUT_DIR, cfg["icm"]["data_file"]) if ICM_ENABLED else ""
ICM_ACTIVE_URL = cfg["icm"]["active_query_url"] if ICM_ENABLED else ""
ICM_RESOLVED_URL = cfg["icm"]["resolved_query_url"] if ICM_ENABLED else ""

EM_ENABLED = cfg["em_layout"]["enabled"]
EM_PAIRS = cfg["em_layout"]["pairs"]

SPRINT_TAB_PATTERN = cfg["excel"]["sprint_tab_pattern"]
DEPENDENCIES_TAB = cfg["excel"]["dependencies_tab"]
EM_TAB = cfg["excel"]["em_tab"] if EM_ENABLED else None

PRIORITY_ORDER = cfg["priority_order"]
# STATUS_ORDER and ROLLED_OUT_STATUSES are read from Excel Config tab below

# ============================================================
# STEP 1: Read Excel data
# ============================================================
print("=== Reading shiproom data ===")
try:
    excel = win32com.client.GetActiveObject("Excel.Application")
    wb = None
    for i in range(1, excel.Workbooks.Count + 1):
        if excel.Workbooks(i).Name == os.path.basename(EXCEL_PATH):
            wb = excel.Workbooks(i)
            break
    if not wb:
        wb = excel.Workbooks.Open(EXCEL_PATH)
except:
    excel = win32com.client.DispatchEx("Excel.Application")
    wb = excel.Workbooks.Open(EXCEL_PATH)

# Read StatusList and PriorityList from the Config tab in Excel
STATUS_ORDER = []
try:
    config_sheet = wb.Sheets("Config")
    for r in range(2, 500):
        val = config_sheet.Cells(r, 1).Value  # Column A = StatusList
        if not val:
            break
        STATUS_ORDER.append(str(val).strip())
    print(f"  StatusList from Excel Config tab: {len(STATUS_ORDER)} statuses")
except Exception as ex:
    print(f"  WARNING: Could not read StatusList from Config tab: {ex}")
    STATUS_ORDER = ["Not Started", "In PM Planning", "PM Spec Complete", "In Design",
        "Design Complete", "In Engg", "Engg POC complete", "Engg Complete",
        "Rolled Out to DF", "Rolled Out to MSIT", "Rolled Out to Prod"]

# Derive rolled-out statuses: only statuses containing "Rolled Out"
ROLLED_OUT_STATUSES = {s for s in STATUS_ORDER if "rolled out" in s.lower()}
print(f"  Rolled-out statuses (auto-detected): {ROLLED_OUT_STATUSES}")

# Dynamic month tab from config pattern
current_month = datetime.now().strftime("%b")
current_year = datetime.now().strftime("%Y")
_month_long = datetime.now().strftime("%B")
_month_num_tab = datetime.now().strftime("%m")
sprint_tab_name = SPRINT_TAB_PATTERN.format(
    month_short=current_month, year=current_year,
    month_long=_month_long, month_num=_month_num_tab
)

tab_found = False
all_tabs = []
for s in range(1, wb.Sheets.Count + 1):
    all_tabs.append(wb.Sheets(s).Name)
    if wb.Sheets(s).Name == sprint_tab_name:
        tab_found = True

if not tab_found:
    print(f"\n  ERROR: Tab '{sprint_tab_name}' not found.")
    print(f"  Available tabs: {all_tabs}")
    print(f"\n  Please create and fill in '{sprint_tab_name}' first, then re-run.")
    sys.exit(1)

print(f"  Using sprint tab: '{sprint_tab_name}'")
sheet = wb.Sheets(sprint_tab_name)

# Build header name → column number mapping (resilient to column shifts)
col_map = {}
for c in range(1, 30):
    h = sheet.Cells(1, c).Value
    if h:
        col_map[str(h).strip().lower()] = c

def col(name):
    """Look up column number by header name (case-insensitive, partial match)."""
    name_l = name.lower()
    # Exact match first
    if name_l in col_map:
        return col_map[name_l]
    # Partial match (header contains name)
    for h, c in col_map.items():
        if name_l in h:
            return c
    return None

def cell_str(row, col_num):
    """Read cell as stripped string, return empty string if None."""
    if not col_num:
        return ""
    return str(sheet.Cells(row, col_num).Value or "").strip()

# Map known headers
COL_ID = col("item id")
COL_EPIC = col("epic")
COL_WORK_ITEM = col("work item")
COL_WORK_TYPE = col("work type")
COL_PRIORITY = col("priority")
COL_PM = col("pm owner")
COL_DEV = col("dev owner")
COL_DESIGN = col("design owner")
COL_START = col("sprint start")
COL_CURRENT = col("current status")
COL_TARGET = col("target status")
COL_RELEASE = col("release to prod")
COL_RELEASE_ACTUAL = col("release to prod month (actual)") or col("release to prod")
COL_RELEASE_COMMITTED = col("release to prod month (committed)")
COL_BLOCKER = col("blocker") # matches "Blocker", "Blockers/Risks", etc.
COL_BLOCKER_DETAILS = col("blocker detail")
COL_PATH_GREEN = col("path to green")
COL_ETA_UNBLOCK = col("eta to unblock")
COL_ASK = col("ask")
COL_PITCH = col("feature pitch")

print(f"  Column mapping: ID={COL_ID}, Epic={COL_EPIC}, WorkItem={COL_WORK_ITEM}")
print(f"    Blocker={COL_BLOCKER}, Details={COL_BLOCKER_DETAILS}, PathGreen={COL_PATH_GREEN}, ETA={COL_ETA_UNBLOCK}")
print(f"    ReleaseActual={COL_RELEASE_ACTUAL}, ReleaseCommitted={COL_RELEASE_COMMITTED}, Pitch={COL_PITCH}")

sprint_items = []
for r in range(2, 500):
    item_id = sheet.Cells(r, COL_ID).Value if COL_ID else None
    work_item = sheet.Cells(r, COL_WORK_ITEM).Value if COL_WORK_ITEM else None
    if not item_id and not work_item:
        break
    if not item_id or not work_item:
        continue
    sprint_items.append({
        "id": cell_str(r, COL_ID),
        "epic": cell_str(r, COL_EPIC),
        "work_item": cell_str(r, COL_WORK_ITEM),
        "work_type": cell_str(r, COL_WORK_TYPE),
        "priority": cell_str(r, COL_PRIORITY),
        "pm_owner": cell_str(r, COL_PM),
        "dev_owner": cell_str(r, COL_DEV),
        "design_owner": cell_str(r, COL_DESIGN),
        "sprint_start_status": cell_str(r, COL_START),
        "current_status": cell_str(r, COL_CURRENT),
        "target_status": cell_str(r, COL_TARGET),
        "release_actual": cell_str(r, COL_RELEASE_ACTUAL),
        "release_committed": cell_str(r, COL_RELEASE_COMMITTED),
        "blocker": cell_str(r, COL_BLOCKER),
        "blocker_details": cell_str(r, COL_BLOCKER_DETAILS),
        "path_to_green": cell_str(r, COL_PATH_GREEN),
        "eta_unblock": cell_str(r, COL_ETA_UNBLOCK),
        "ask": cell_str(r, COL_ASK),
        "pitch": cell_str(r, COL_PITCH),
    })

# Dependencies (optional)
dependencies = []
if DEPENDENCIES_TAB:
    try:
        dep_sheet = wb.Sheets(DEPENDENCIES_TAB)
        for r in range(2, 500):
            dep_id = dep_sheet.Cells(r, 1).Value
            desc = dep_sheet.Cells(r, 4).Value
            if not dep_id and not desc:
                break
            if not dep_id or not desc:
                continue
            dependencies.append({
                "work_stream": str(dep_sheet.Cells(r, 2).Value or "").strip(),
                "work_item": str(dep_sheet.Cells(r, 3).Value or "").strip(),
                "description": str(desc).strip(),
                "partner_team": str(dep_sheet.Cells(r, 6).Value or "").strip(),
                "date_needed": str(dep_sheet.Cells(r, 9).Value or "").strip()[:10],
                "status": str(dep_sheet.Cells(r, 10).Value or "").strip(),
                "risk_level": str(dep_sheet.Cells(r, 13).Value or "").strip(),
                "impact": str(dep_sheet.Cells(r, 14).Value or "").strip(),
                "mitigation": str(dep_sheet.Cells(r, 15).Value or "").strip(),
                "owner": str(dep_sheet.Cells(r, 16).Value or "").strip(),
                "notes": str(dep_sheet.Cells(r, 17).Value or "").strip(),
            })
    except Exception as ex:
        print(f"  WARNING: Could not read dependencies tab '{DEPENDENCIES_TAB}': {ex}")
else:
    print("  Dependencies tab: skipped (not configured)")

# ============================================================
# STEP 1b: Read EM tab → build IC-to-EM mapping
# ============================================================
em_mapping = {}  # IC name → EM name
if EM_ENABLED and EM_TAB:
    try:
        em_sheet = wb.Sheets(EM_TAB)
        for pair in EM_PAIRS:
            em_name_row = pair.get("em_name_row", 2)
            em_col = pair["em_col"]
            direct_col = pair["directs_col"]
            em_name = str(em_sheet.Cells(em_name_row, em_col).Value or "").strip()
            if not em_name:
                continue
            for r in range(2, 500):
                direct = em_sheet.Cells(r, direct_col).Value
                if not direct:
                    break
                em_mapping[str(direct).strip()] = em_name
        print(f"  EM mapping: {len(em_mapping)} ICs loaded ({', '.join(sorted(set(em_mapping.values())))})")
    except Exception as ex:
        print(f"  WARNING: Could not read EM tab '{EM_TAB}': {ex}")
        print(f"  Available tabs: {all_tabs}")
else:
    print("  EM tab: skipped (not configured)")

# Derive em_owner for each sprint item
# Dev Owner can be an IC name (lookup in em_mapping) OR an EM name directly
em_name_set = set(em_mapping.values())
for item in sprint_items:
    dev = item["dev_owner"]
    if dev in em_mapping:
        item["em_owner"] = em_mapping[dev]
    elif dev in em_name_set:
        item["em_owner"] = dev  # Dev Owner IS the EM
    else:
        item["em_owner"] = ""

print(f"  Sprint items: {len(sprint_items)}, Dependencies: {len(dependencies)}")

# ============================================================
# STEP 2: Compute KPIs
# ============================================================
total = len(sprint_items)
sc = Counter(i["current_status"] for i in sprint_items)
in_engg = sum(sc.get(s, 0) for s in ["In Engg", "Engg Complete"])
in_rollout = sum(sc.get(s, 0) for s in ["Rolled Out to DF", "Rolled Out to MSIT"])
shipped = sc.get("Rolled Out to Prod", 0)

has_blocker = [i for i in sprint_items if i["blocker"] and i["blocker"] not in ("None", "", "none")]
blocked_count = len(has_blocker)

rollout_df_msit = [i for i in sprint_items if i["target_status"] in ("Rolled Out to DF", "Rolled Out to MSIT")]
rollout_prod = [i for i in sprint_items if i["target_status"] == "Rolled Out to Prod"]
asks = [i for i in sprint_items if i["ask"] and i["ask"].strip()]

blocked_pct = (blocked_count / total * 100) if total > 0 else 0
if blocked_pct > 50:
    sprint_health = "OFF TRACK"
    health_color = "#D13438"
    health_glow = "rgba(209,52,56,0.35)"
elif blocked_pct > 25:
    sprint_health = "AT RISK"
    health_color = "#FF8C00"
    health_glow = "rgba(255,140,0,0.35)"
else:
    sprint_health = "ON TRACK"
    health_color = "#107C10"
    health_glow = "rgba(16,124,16,0.35)"

health_explanation = f"{blocked_count} of {total} items ({blocked_pct:.0f}%) blocked or at risk"
print(f"  Sprint Health: {sprint_health} ({blocked_pct:.0f}% blocked/at-risk)")

# Sprint progress — mutually exclusive buckets
wins_items = []  # Met or exceeded target
moved_fwd_items = []  # Moved forward from start but not yet at target
no_move_items = []  # No movement from sprint start

for i in sprint_items:
    cur = i["current_status"]
    tgt = i["target_status"]
    start = i["sprint_start_status"]
    if cur not in STATUS_ORDER or tgt not in STATUS_ORDER or start not in STATUS_ORDER:
        no_move_items.append(i)
        continue
    cur_idx = STATUS_ORDER.index(cur)
    tgt_idx = STATUS_ORDER.index(tgt)
    start_idx = STATUS_ORDER.index(start)
    if cur_idx >= tgt_idx:
        wins_items.append(i)
    elif cur_idx > start_idx:
        moved_fwd_items.append(i)
    else:
        no_move_items.append(i)

wins_count = len(wins_items)
moved_forward = len(moved_fwd_items)
stayed_same = len(no_move_items)
progress_pct = int(wins_count / total * 100) if total > 0 else 0
print(f"  Wins: {wins_count} met target | Moved forward: {moved_forward} | No movement: {stayed_same} | Total: {wins_count+moved_forward+stayed_same}")

# ============================================================
# STEP 3: ADO bug queries (conditional on config)
# ============================================================
now = datetime.now()
_month_num = now.strftime("%m")
_month_short = now.strftime("%b")
_year = now.strftime("%Y")
_half = "H1" if now.month <= 6 else "H2"
_quarter = f"Q{(now.month - 1) // 3 + 1}"
_cy = f"CY{_year[2:]}"

p1_count = 0; p2_count = 0; resolved_count = 0; opened_14d_count = 0
p1_ids = []; p2_ids = []; resolved_ids = []
iteration_path = ""

if ADO_ENABLED:
    iteration_path = ADO_ITERATION_TEMPLATE.format(
        project=ADO_PROJECT, cy=_cy, half=_half, quarter=_quarter,
        year=_year, month_num=_month_num, month_short=_month_short
    )
    print(f"  Bug iteration path: {iteration_path}")

    try:
        _h = get_ado_auth_headers()
        _ado_url = f"{ADO_API_BASE}/wit/wiql?api-version=7.1"

        # P1 bugs
        _wiql_p1 = {"query": f"SELECT [System.Id] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] = 1 AND [System.IterationPath] = '{iteration_path}' AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] NOT IN ('Done', 'Removed')"}
        _r = requests.post(_ado_url, headers=_h, json=_wiql_p1)
        if _r.status_code == 200:
            p1_ids = [str(wi["id"]) for wi in _r.json().get("workItems", [])]
            p1_count = len(p1_ids)

        # P2 bugs
        _wiql_p2 = {"query": f"SELECT [System.Id] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] = 2 AND [System.IterationPath] = '{iteration_path}' AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] NOT IN ('Done', 'Removed')"}
        _r = requests.post(_ado_url, headers=_h, json=_wiql_p2)
        if _r.status_code == 200:
            p2_ids = [str(wi["id"]) for wi in _r.json().get("workItems", [])]
            p2_count = len(p2_ids)

        # Resolved bugs (14d) — with retry
        for _attempt in range(3):
            _wiql_res = {"query": f"SELECT [System.Id] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] IN (1, 2) AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] IN ('Done', 'Removed') AND [Microsoft.VSTS.Common.ClosedDate] >= @StartOfDay('-14d')"}
            _r = requests.post(_ado_url, headers=_h, json=_wiql_res)
            if _r.status_code == 200:
                resolved_ids = [str(wi["id"]) for wi in _r.json().get("workItems", [])]
                resolved_count = len(resolved_ids)
                if resolved_count > 0:
                    break
            import time; time.sleep(1)

        # Opened P1+P2 bugs in last 14 days
        _wiql_opened = {"query": f"SELECT [System.Id] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] IN (1, 2) AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.CreatedDate] >= @StartOfDay('-14d') AND [System.State] NOT IN ('Done', 'Removed')"}
        _r = requests.post(_ado_url, headers=_h, json=_wiql_opened)
        if _r.status_code == 200:
            opened_14d_count = len(_r.json().get("workItems", []))

    except Exception as ex:
        print(f"  Bug query error: {ex}")
        if BUG_DATA_PATH:
            try:
                with open(BUG_DATA_PATH, "r", encoding="utf-8") as f:
                    bug_data = json.load(f)
                p1_count = bug_data["summary"]["p1_count"]
                p2_count = bug_data["summary"]["p2_count"]
            except: pass

    print(f"  Bugs - P1:{p1_count}, P2:{p2_count}, Resolved(14d):{resolved_count}, Opened(14d):{opened_14d_count}")
else:
    print("  ADO bugs: skipped (not configured)")

# ============================================================
# STEP 3b: Fetch bug assignees → build per-EM aggregation
# ============================================================
def fetch_bug_assignees(bug_ids, headers):
    """Fetch System.AssignedTo for a list of bug IDs via batch API."""
    assignees = {}  # id → display name
    if not bug_ids:
        return assignees
    # Batch in groups of 200
    for i in range(0, len(bug_ids), 200):
        batch = bug_ids[i:i+200]
        ids_str = ",".join(batch)
        url = f"{ADO_API_BASE}/wit/workitems?ids={ids_str}&fields=System.Id,System.AssignedTo,Microsoft.VSTS.Common.Priority&api-version=7.1"
        try:
            r = requests.get(url, headers=headers)
            if r.status_code == 200:
                for wi in r.json().get("value", []):
                    assigned = wi.get("fields", {}).get("System.AssignedTo", {})
                    name = assigned.get("displayName", "") if isinstance(assigned, dict) else str(assigned or "")
                    assignees[str(wi["id"])] = name
        except:
            pass
    return assignees

# Build per-EM bug aggregation
bugs_by_em = defaultdict(lambda: {"p1": 0, "p2": 0, "resolved": 0})

if ADO_ENABLED:
    try:
        _h_get = {k: v for k, v in _h.items() if k == "Authorization"}

        # Build a first-name lookup for EM resolution from ADO display names
        first_name_to_em = {}
        for ic_name, em_name in em_mapping.items():
            first = ic_name.split()[0].strip().lower()
            first_name_to_em[first] = em_name
            first_name_to_em[ic_name.strip().lower()] = em_name

        def resolve_em(ado_display_name):
            """Resolve an ADO display name to an EM via exact match, then first-name match."""
            if not ado_display_name:
                return "Unassigned"
            name = ado_display_name.strip()
            if name in em_mapping:
                return em_mapping[name]
            name_lower = name.lower()
            for ic, em in em_mapping.items():
                if ic.lower() == name_lower:
                    return em
            first = name.split()[0].strip().lower()
            if first in first_name_to_em:
                return first_name_to_em[first]
            return "Unassigned"

        if p1_ids or p2_ids:
            open_assignees = fetch_bug_assignees(p1_ids + p2_ids, _h_get)
            for bug_id in p1_ids:
                name = open_assignees.get(bug_id, "")
                em = resolve_em(name)
                bugs_by_em[em]["p1"] += 1
            for bug_id in p2_ids:
                name = open_assignees.get(bug_id, "")
                em = resolve_em(name)
                bugs_by_em[em]["p2"] += 1

        if resolved_ids:
            resolved_assignees = fetch_bug_assignees(resolved_ids, _h_get)
            for bug_id in resolved_ids:
                name = resolved_assignees.get(bug_id, "")
                em = resolve_em(name)
                bugs_by_em[em]["resolved"] += 1

        print(f"  Bug assignees fetched. EMs with bugs: {dict(bugs_by_em)}")
    except Exception as ex:
        print(f"  Bug assignee fetch error: {ex}")

# Build bug query URLs for hyperlinks (only if ADO enabled)
P1_URL = ""
P2_URL = ""
RESOLVED_URL = ""
OPENED_URL = ""
if ADO_ENABLED:
    _wiql_p1_raw = f"SELECT [System.Id], [System.Title] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] = 1 AND [System.IterationPath] = '{iteration_path}' AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] NOT IN ('Done', 'Removed')"
    _wiql_p2_raw = f"SELECT [System.Id], [System.Title] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] = 2 AND [System.IterationPath] = '{iteration_path}' AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] NOT IN ('Done', 'Removed')"
    _wiql_resolved_raw = f"SELECT [System.Id], [System.Title] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] IN (1, 2) AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.State] IN ('Done', 'Removed') AND [Microsoft.VSTS.Common.ClosedDate] >= @StartOfDay('-14d')"
    _wiql_opened_raw = f"SELECT [System.Id], [System.Title] FROM workitems WHERE [System.WorkItemType] = 'Bug' AND [Microsoft.VSTS.Common.Priority] IN (1, 2) AND [System.AreaPath] = '{ADO_AREA_PATH}' AND [System.CreatedDate] >= @StartOfDay('-14d') AND [System.State] NOT IN ('Done', 'Removed')"
    _ado_query_base = f"https://dev.azure.com/{quote(ADO_ORG)}/{quote(ADO_PROJECT)}/_queries/query/?wiql="
    P1_URL = f"{_ado_query_base}{quote(_wiql_p1_raw)}"
    P2_URL = f"{_ado_query_base}{quote(_wiql_p2_raw)}"
    RESOLVED_URL = f"{_ado_query_base}{quote(_wiql_resolved_raw)}"
    OPENED_URL = f"{_ado_query_base}{quote(_wiql_opened_raw)}"

# ============================================================
# STEP 3c: Read ICM data from JSON (exported by _fetch_icm.py)
# ============================================================
icm_active = []
icm_resolved = []
icm_active_count = 0
icm_resolved_count = 0
if ICM_ENABLED:
    # Always fetch fresh ICM data via _fetch_icm.py
    import subprocess
    _icm_script = os.path.join(os.path.dirname(EXCEL_PATH), "_fetch_icm.py")
    if os.path.exists(_icm_script):
        print("  ICM: fetching fresh data (launching browser)...")
        _icm_result = subprocess.run(
            [sys.executable, _icm_script, "--config", os.path.abspath(args.config)],
            capture_output=True, text=True, timeout=180
        )
        if _icm_result.returncode != 0:
            print(f"  ICM fetch warning: {_icm_result.stderr.strip()[-200:]}")
        else:
            print(f"  ICM fetch complete")
    else:
        print(f"  WARNING: _fetch_icm.py not found at {_icm_script}")

    try:
        with open(ICM_DATA_PATH, "r", encoding="utf-8") as f:
            icm_data = json.load(f)
        icm_active = icm_data.get("active", [])
        icm_resolved = icm_data.get("resolved", [])
        icm_active_count = len(icm_active)
        icm_resolved_count = len(icm_resolved)
        print(f"  ICM data: {icm_active_count} active, {icm_resolved_count} resolved")
    except Exception as ex:
        print(f"  ICM data error: {ex}")
else:
    print("  ICM: skipped (not configured)")

# ============================================================
# STEP 4: Group by epic + compute per-EM resource utilization
# ============================================================
epic_order = []; epic_groups = {}
for item in sprint_items:
    ep = item["epic"]
    if ep not in epic_groups: epic_order.append(ep); epic_groups[ep] = []
    epic_groups[ep].append(item)
for ep in epic_order:
    epic_groups[ep].sort(key=lambda x: PRIORITY_ORDER.get(x["priority"], 99))

def count_statuses(items):
    c = {"Not Started":0,"In PM/Design":0,"In Engg":0,"In Rollout":0,"Shipped":0,"Blocked":0}
    for i in items:
        s=i["current_status"]; b=i.get("blocker","")
        if b and b not in ("None","","none"): c["Blocked"]+=1
        elif s in ("Not Started",): c["Not Started"]+=1
        elif s in ("In PM Planning","PM Spec Complete","In Design","Design Complete"): c["In PM/Design"]+=1
        elif s in ("In Engg","Engg Complete"): c["In Engg"]+=1
        elif s in ("Rolled Out to DF","Rolled Out to MSIT"): c["In Rollout"]+=1
        elif s=="Rolled Out to Prod": c["Shipped"]+=1
    return c

# Per-EM resource utilization from feature items
em_names = sorted(set(em_mapping.values())) if em_mapping else []
em_feature_data = {}
for em in em_names:
    em_items = [i for i in sprint_items if i.get("em_owner") == em]
    em_cnts = count_statuses(em_items)
    rollout_count = sum(1 for i in em_items if i["target_status"] in ("Rolled Out to DF", "Rolled Out to MSIT", "Rolled Out to Prod"))
    em_feature_data[em] = {
        "total": len(em_items),
        "rollout": rollout_count,
        "in_engg": em_cnts["In Engg"],
        "shipped": em_cnts["Shipped"],
        "blocked": em_cnts["Blocked"],
    }

print(f"\n=== Data summary ===")
print(f"  Rollout DF/MSIT: {len(rollout_df_msit)}, Rollout Prod: {len(rollout_prod)}")
print(f"  Asks: {len(asks)}, Dependencies: {len(dependencies)}")
print(f"  Epics: {len(epic_order)}")
for em in em_names:
    fd = em_feature_data[em]
    bd = bugs_by_em.get(em, {"p1":0,"p2":0,"resolved":0})
    print(f"  {em}: {fd['total']} features, {bd['p1']} P1 bugs, {bd['p2']} P2 bugs, {bd['resolved']} resolved")

# ============================================================
# STEP 5: HTML generation helpers
# ============================================================
def e(text):
    """HTML-escape text."""
    return escape(str(text)) if text else ""

STATUS_CLASS_MAP = {
    "Not Started": "status-not-started",
    "In PM Planning": "status-pm-planning",
    "PM Spec Complete": "status-pm-spec",
    "In Design": "status-in-design",
    "Design Complete": "status-design-complete",
    "In Engg": "status-in-engg",
    "Engg Complete": "status-engg-complete",
    "Engg POC complete": "status-engg-poc-complete",
    "On Hold": "status-on-hold",
    "Blocked": "status-blocked",
    "Deprioritised": "status-deprioritised",
    "Rolled Out to DF": "status-rolled-df",
    "Rolled Out to MSIT": "status-rolled-msit",
    "Rolled Out to Prod": "status-rolled-prod",
}

def status_pill(status):
    cls = STATUS_CLASS_MAP.get(status, "")
    return f'<span class="status-pill {cls}">{e(status)}</span>'

def has_risk(item):
    return item["blocker"] and item["blocker"] not in ("None", "", "none")

def is_win(item):
    return item["current_status"] in ROLLED_OUT_STATUSES and not has_risk(item)

def row_class(item):
    if has_risk(item): return ' class="row-blocked"'
    if is_win(item): return ' class="row-shipped"'
    return ''

WARN_SVG = '<svg class="blocker-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'

def blocker_html(item):
    if not has_risk(item):
        return ''
    blocker = item["blocker"]
    details = item.get("blocker_details", "")
    text = f"{blocker}: {details}" if details else blocker
    is_at_risk = "at risk" in blocker.lower() or "at risk" in text.lower()
    cls = "blocker-text at-risk" if is_at_risk else "blocker-text"
    return f'<div class="{cls}">{WARN_SVG} {e(text)}</div>'

def blocker_html_short(item):
    if not has_risk(item):
        return ''
    blocker = item["blocker"]
    is_at_risk = "at risk" in blocker.lower()
    cls = "blocker-text at-risk" if is_at_risk else "blocker-text"
    return f'<span class="{cls}">{e(blocker)}</span>'

def priority_html(priority):
    return f'<span class="priority-badge priority-{e(priority)}">{e(priority)}</span>'

def work_item_display(item):
    wi = e(item["work_item"])
    if item.get("work_type", "").lower() == "customer feedback":
        wi += ' <span class="customer-pill">Customer</span>'
    pitch = item.get("pitch", "")
    if pitch:
        wi += f'<div class="wi-pitch" onclick="this.classList.toggle(\'open\');event.stopPropagation()"><span class="wi-pitch-toggle">&#9662; Details</span><div class="wi-pitch-text">{e(pitch)}</div></div>'
    return wi

def path_to_green_cell(item):
    """Format Path to Green as [ETA] description, only for blocked/at-risk items."""
    if not has_risk(item):
        return ''
    eta = item.get("eta_unblock", "")
    path = item.get("path_to_green", "")
    if not eta and not path:
        return '<span style="color:var(--gray-400);font-size:11px;font-style:italic;">Not provided</span>'
    parts = []
    if eta:
        parts.append(f'<strong style="color:var(--navy);">[{e(eta)}]</strong>')
    if path:
        parts.append(e(path))
    return f'<span style="font-size:12px;line-height:1.4;">{" ".join(parts)}</span>'

# ============================================================
# STEP 6: Build grouped collapsible section (items grouped by epic)
# ============================================================
def owners_cell(item):
    """PM / EM (or Dev) / Design owners in one cell with delimiters."""
    parts = []
    if item.get("pm_owner"): parts.append(f'<strong>PM:</strong> {e(item["pm_owner"])}')
    if EM_ENABLED and item.get("em_owner"):
        parts.append(f'<strong>EM:</strong> {e(item["em_owner"])}')
    elif not EM_ENABLED and item.get("dev_owner"):
        parts.append(f'<strong>Dev:</strong> {e(item["dev_owner"])}')
    if item.get("design_owner"): parts.append(f'<strong>Des:</strong> {e(item["design_owner"])}')
    return '<span style="font-size:11px;line-height:1.6;">' + '<br>'.join(parts) + '</span>' if parts else ''

def eta_cells(item):
    """Committed and Actual ETA columns."""
    committed = e(item.get("release_committed", ""))
    actual = e(item.get("release_actual", ""))
    # Highlight if actual differs from committed (slip)
    slip_cls = ""
    if committed and actual and committed != actual:
        slip_cls = ' style="color:var(--red);font-weight:600;"'
    return (f'<td>{committed}</td>',
            f'<td{slip_cls}>{actual}</td>')

def _item_row(item):
    committed_td, actual_td = eta_cells(item)
    return f'''              <tr{row_class(item)}>
                <td>{work_item_display(item)}</td>
                <td>{owners_cell(item)}</td>
                <td>{status_pill(item["current_status"])}</td>
                <td>{status_pill(item["target_status"])}</td>
                {committed_td}
                {actual_td}
                <td>{blocker_html(item)}</td>
                <td>{path_to_green_cell(item)}</td>
              </tr>'''

TABLE_HEADER = '''          <thead>
            <tr>
              <th style="width:18%">Work Item</th>
              <th style="width:10%">Owners</th>
              <th style="width:11%">Current Status</th>
              <th style="width:11%">Target Status</th>
              <th style="width:7%">Committed ETA</th>
              <th style="width:7%">Actual ETA</th>
              <th style="width:20%">Blocker / Risk</th>
              <th style="width:16%">Path to Green</th>
            </tr>
          </thead>'''

def grouped_section_html(items, prefix):
    """Group items by epic into collapsible accordion sections."""
    # Group preserving order
    order = []
    groups = {}
    for item in items:
        ep = item["epic"]
        if ep not in groups:
            order.append(ep)
            groups[ep] = []
        groups[ep].append(item)

    sections = []
    for idx, ep in enumerate(order):
        ep_items = groups[ep]
        cnts = count_statuses(ep_items)
        rows_html = '\n'.join(_item_row(i) for i in ep_items)
        gid = f"{prefix}-{idx}"
        icon_cls = f"epic-icon-{idx % 3}"
        icon = EPIC_ICONS[idx % len(EPIC_ICONS)]

        # Mini KPIs for the header
        mkpis = []
        if cnts["In Engg"] > 0:
            mkpis.append(f'<span class="mini-kpi mini-kpi-engg">Engg {cnts["In Engg"]}</span>')
        if cnts["In Rollout"] > 0:
            mkpis.append(f'<span class="mini-kpi mini-kpi-rollout">Rollout {cnts["In Rollout"]}</span>')
        if cnts["Shipped"] > 0:
            mkpis.append(f'<span class="mini-kpi mini-kpi-shipped">Shipped {cnts["Shipped"]}</span>')
        if cnts["Blocked"] > 0:
            mkpis.append(f'<span class="mini-kpi mini-kpi-blocked">Blocked/At Risk {cnts["Blocked"]}</span>')
        mkpis_html = ' '.join(mkpis)

        open_cls = " open" if idx == 0 else ""
        sections.append(f'''      <div class="epic-card{open_cls}" id="{gid}">
        <div class="epic-card-header" onclick="toggleEpic('{gid}')">
          <div class="epic-card-left">
            <div class="epic-icon {icon_cls}">{icon}</div>
            <div>
              <div class="epic-name">{e(ep)}</div>
              <div class="epic-item-count">{len(ep_items)} items</div>
            </div>
          </div>
          <div class="epic-card-right">
            <div class="epic-mini-kpis">{mkpis_html}</div>
            <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
          </div>
        </div>
        <div class="epic-card-body">
          <div class="epic-card-body-inner">
            <table>
{TABLE_HEADER}
              <tbody>
{rows_html}
              </tbody>
            </table>
          </div>
        </div>
      </div>''')

    return '\n'.join(sections)

# Keep flat row builder for blocked items table (reused internally)
def rollout_table_rows(items):
    return '\n'.join(_item_row(i) for i in items)

# ============================================================
# STEP 7: Build asks table rows
# ============================================================
def asks_table_rows(items):
    rows = []
    for item in items:
        is_blocked = "blocked" in item.get("blocker", "").lower()
        color = "var(--red)" if is_blocked else "var(--orange)"
        rows.append(f'''              <tr>
                <td>{e(item["epic"])}</td>
                <td>{e(item["work_item"])}</td>
                <td><span style="font-weight:600;color:{color};">{e(item["ask"])}</span></td>
                <td>{e(item["pm_owner"])}</td>
              </tr>''')
    return '\n'.join(rows)

# ============================================================
# STEP 8: Build dependency table rows
# ============================================================
def dep_table_rows(deps):
    rows = []
    for dep in deps:
        risk = dep.get("risk_level", "").lower()
        cls = ' class="row-blocked"' if "high" in risk or "critical" in risk else ''
        rows.append(f'''              <tr{cls}>
                <td>{e(dep["work_stream"])}</td>
                <td>{e(dep["work_item"])}</td>
                <td>{e(dep["impact"])}</td>
                <td>{e(dep["mitigation"])}</td>
                <td>{e(dep["owner"])}</td>
              </tr>''')
    return '\n'.join(rows)

# ============================================================
# STEP 9: Build epic tracker cards (with EM + Design Owner, renamed pill)
# ============================================================
EPIC_ICONS = ["&#128240;", "&#128196;", "&#128247;", "&#128161;", "&#128295;", "&#128202;", "&#128640;", "&#127919;"]

def epic_card_html(idx, epic_name, items):
    cnts = count_statuses(items)
    icon = EPIC_ICONS[idx % len(EPIC_ICONS)]
    icon_cls = f"epic-icon-{idx % 3}"

    # Mini KPI badges - only show non-zero; renamed "Blocked" → "Blocked/At Risk"
    mini_kpis = []
    if cnts["In PM/Design"] > 0:
        mini_kpis.append(f'<span class="mini-kpi" style="background:#F3E5F5;color:var(--purple);">PM/Design {cnts["In PM/Design"]}</span>')
    if cnts["In Engg"] > 0:
        mini_kpis.append(f'<span class="mini-kpi mini-kpi-engg">Engg {cnts["In Engg"]}</span>')
    if cnts["In Rollout"] > 0:
        mini_kpis.append(f'<span class="mini-kpi mini-kpi-rollout">Rollout {cnts["In Rollout"]}</span>')
    if cnts["Shipped"] > 0:
        mini_kpis.append(f'<span class="mini-kpi mini-kpi-shipped">Shipped {cnts["Shipped"]}</span>')
    if cnts["Blocked"] > 0:
        mini_kpis.append(f'<span class="mini-kpi mini-kpi-blocked">Blocked/At Risk {cnts["Blocked"]}</span>')
    mini_kpis_html = '\n            '.join(mini_kpis)

    # Table rows
    item_rows = []
    for item in items:
        committed_td, actual_td = eta_cells(item)
        item_rows.append(f'''              <tr{row_class(item)}>
                <td>{work_item_display(item)}</td>
                <td>{owners_cell(item)}</td>
                <td>{status_pill(item["current_status"])}</td>
                <td>{status_pill(item["target_status"])}</td>
                {committed_td}
                {actual_td}
                <td>{priority_html(item["priority"])}</td>
                <td>{blocker_html(item)}</td>
                <td>{path_to_green_cell(item)}</td>
              </tr>''')
    item_rows_html = '\n'.join(item_rows)

    return f'''    <div class="epic-card" id="epic-{idx}">
      <div class="epic-card-header" onclick="toggleEpic('epic-{idx}')">
        <div class="epic-card-left">
          <div class="epic-icon {icon_cls}">{icon}</div>
          <div>
            <div class="epic-name">{e(epic_name)}</div>
            <div class="epic-item-count">{len(items)} items</div>
          </div>
        </div>
        <div class="epic-card-right">
          <div class="epic-mini-kpis">
            {mini_kpis_html}
          </div>
          <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
        </div>
      </div>
      <div class="epic-card-body">
        <div class="epic-card-body-inner">
          <table class="epic-table">
            <thead>
              <tr>
                <th style="width:16%">Work Item</th>
                <th style="width:9%">Owners</th>
                <th style="width:10%">Current</th>
                <th style="width:10%">Target</th>
                <th style="width:6%">Committed</th>
                <th style="width:6%">Actual</th>
                <th style="width:5%">Priority</th>
                <th style="width:20%">Blocker/Risks</th>
                <th style="width:14%">Path to Green</th>
              </tr>
            </thead>
            <tbody>
{item_rows_html}
            </tbody>
          </table>
        </div>
      </div>
    </div>'''

# ============================================================
# STEP 9b: Build Resource Utilization cards HTML
# ============================================================
def resource_card_html(em_name, feature_data, bug_data):
    fd = feature_data
    bd = bug_data
    return f'''    <div class="resource-card">
      <div class="resource-card-header">
        <div class="resource-card-name">{e(em_name)}&rsquo;s Team</div>
      </div>
      <div class="resource-card-kpis">
        <div class="resource-kpi" style="--rk-color: var(--blue);">
          <div class="resource-kpi-value" data-target="{fd['total']}">{fd['total']}</div>
          <div class="resource-kpi-label">Features</div>
        </div>
        <div class="resource-kpi" style="--rk-color: var(--teal);">
          <div class="resource-kpi-value" data-target="{fd['rollout']}">{fd['rollout']}</div>
          <div class="resource-kpi-label">Rollout Target</div>
        </div>
        <div class="resource-kpi-separator"></div>
        <div class="resource-kpi" style="--rk-color: var(--red);">
          <div class="resource-kpi-value" data-target="{bd['p1']}">{bd['p1']}</div>
          <div class="resource-kpi-label">Open P1</div>
        </div>
        <div class="resource-kpi" style="--rk-color: var(--orange);">
          <div class="resource-kpi-value" data-target="{bd['p2']}">{bd['p2']}</div>
          <div class="resource-kpi-label">Open P2</div>
        </div>
        <div class="resource-kpi" style="--rk-color: var(--green);">
          <div class="resource-kpi-value" data-target="{bd['resolved']}">{bd['resolved']}</div>
          <div class="resource-kpi-label">Closed (14d)</div>
        </div>
      </div>
    </div>'''

resource_cards_html = ""
if em_names:
    cards = []
    for em in em_names:
        fd = em_feature_data.get(em, {"total":0,"rollout":0,"in_engg":0,"shipped":0,"blocked":0})
        bd = bugs_by_em.get(em, {"p1":0,"p2":0,"resolved":0})
        cards.append(resource_card_html(em, fd, bd))
    resource_cards_html = '\n'.join(cards)

# Items slipping committed ETA (actual differs from committed)
slipping_items = [i for i in sprint_items
    if i.get("release_committed") and i.get("release_actual")
    and i["release_committed"].strip() != i["release_actual"].strip()]
print(f"  Slipping items: {len(slipping_items)}")

# Build item lists for KPI card modals
blocked_items_rows = rollout_table_rows(has_blocker)

items_total = sprint_items
items_in_engg = [i for i in sprint_items if i["current_status"] in ("In Engg", "Engg Complete", "Engg POC complete")]
items_in_rollout = [i for i in sprint_items if i["current_status"] in ("Rolled Out to DF", "Rolled Out to MSIT")]
items_shipped = [i for i in sprint_items if i["current_status"] == "Rolled Out to Prod"]

MODAL_TABLE_HEADER = '''          <thead>
            <tr>
              <th style="width:10%">Epic</th>
              <th style="width:16%">Work Item</th>
              <th style="width:9%">Owners</th>
              <th style="width:10%">Current</th>
              <th style="width:10%">Target</th>
              <th style="width:6%">Committed</th>
              <th style="width:6%">Actual</th>
              <th style="width:18%">Blocker / Risk</th>
              <th style="width:15%">Path to Green</th>
            </tr>
          </thead>'''

def _modal_item_row(item):
    committed_td, actual_td = eta_cells(item)
    return f'''              <tr{row_class(item)}>
                <td>{e(item["epic"])}</td>
                <td>{work_item_display(item)}</td>
                <td>{owners_cell(item)}</td>
                <td>{status_pill(item["current_status"])}</td>
                <td>{status_pill(item["target_status"])}</td>
                {committed_td}
                {actual_td}
                <td>{blocker_html(item)}</td>
                <td>{path_to_green_cell(item)}</td>
              </tr>'''

def build_modal(modal_id, title, items):
    rows = '\n'.join(_modal_item_row(i) for i in items)
    return f'''<div class="modal-overlay" id="{modal_id}" onclick="if(event.target===this)this.style.display='none'">
  <div class="modal-content">
    <div class="modal-header">
      <div class="modal-title">{title} <span class="count">{len(items)} items</span></div>
      <button class="modal-close" onclick="document.getElementById('{modal_id}').style.display='none'">&times;</button>
    </div>
    <div class="modal-body">
      <table>
{MODAL_TABLE_HEADER}
        <tbody>
{rows}
        </tbody>
      </table>
    </div>
  </div>
</div>'''

modals_html = '\n'.join([
    build_modal('modal-total', 'All Feature Items', items_total),
    build_modal('modal-rollout', 'In Rollout (DF/MSIT)', items_in_rollout),
    build_modal('modal-shipped', 'Shipped (Preview/Prod)', items_shipped),
    build_modal('modal-wins', 'Met Target or Beyond', wins_items),
])

# ============================================================
# STEP 10: Assemble full HTML
# ============================================================
print("\n=== Generating HTML dashboard ===")

today_str = now.strftime("%d %B %Y").lstrip("0")
month_str = now.strftime("%b %Y")
generated_str = now.strftime("%d %b %Y").lstrip("0")

# Days remaining in sprint (end of current month)
import calendar
_, last_day = calendar.monthrange(now.year, now.month)
sprint_end = now.replace(day=last_day)
days_left = (sprint_end - now).days

epic_cards_html = '\n\n'.join(
    epic_card_html(idx, ep, epic_groups[ep])
    for idx, ep in enumerate(epic_order)
)

# Build conditional ICM KPI cards
icm_cards_html = ""
if ICM_ENABLED:
    icm_cards_html = f'''          <div class="kpi-card" style="--card-accent: var(--red);">
            <a class="kpi-link" href="{ICM_ACTIVE_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{icm_active_count}">{icm_active_count}</div>
              <div class="kpi-label">Active ICMs (14d)</div>
            </a>
          </div>
          <div class="kpi-card" style="--card-accent: var(--green);">
            <a class="kpi-link" href="{ICM_RESOLVED_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{icm_resolved_count}">{icm_resolved_count}</div>
              <div class="kpi-label">Resolved ICMs (14d)</div>
            </a>
          </div>'''

# Build conditional bug KPI group (only if ADO enabled)
icm_summary = f" | ICM: {icm_active_count} active, {icm_resolved_count} resolved" if ICM_ENABLED else ""
bug_kpi_group_html = ""
if ADO_ENABLED:
    bug_kpi_group_html = f'''      <div class="kpi-group kpi-group-bugs collapsed" id="kpi-bugs" onclick="this.classList.toggle('collapsed')">
        <div class="kpi-group-label">Bug Updates{"  &amp; ICMs" if ICM_ENABLED else ""} <span class="kpi-group-summary">({p1_count} P1, {p2_count} P2, {resolved_count} resolved, {opened_14d_count} opened{icm_summary})</span></div>
        <div class="kpi-row">
          <div class="kpi-card" style="--card-accent: var(--red);">
            <a class="kpi-link" href="{P1_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{p1_count}">{p1_count}</div>
              <div class="kpi-label">Open P1</div>
            </a>
          </div>
          <div class="kpi-card" style="--card-accent: var(--orange);">
            <a class="kpi-link" href="{P2_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{p2_count}">{p2_count}</div>
              <div class="kpi-label">Open P2</div>
            </a>
          </div>
          <div class="kpi-card" style="--card-accent: var(--green);">
            <a class="kpi-link" href="{RESOLVED_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{resolved_count}">{resolved_count}</div>
              <div class="kpi-label">Resolved (14d)</div>
            </a>
          </div>
          <div class="kpi-card" style="--card-accent: var(--red);">
            <a class="kpi-link" href="{OPENED_URL}" target="_blank" rel="noopener">
              <div class="kpi-value" data-target="{opened_14d_count}">{opened_14d_count}</div>
              <div class="kpi-label">Open P1, P2 (14d)</div>
            </a>
          </div>
{icm_cards_html}
        </div>
      </div>'''

# Build conditional resource utilization tab button
resource_tab_btn = ""
if EM_ENABLED:
    resource_tab_btn = '''      <button class="tab-btn" data-tab="resource" onclick="switchTab('resource', this)">Resource Utilization</button>'''

# Build conditional dependencies sub-section
deps_subsection_html = ""
if DEPENDENCIES_TAB and dependencies:
    deps_subsection_html = f'''      <div class="sub-section">
        <div class="sub-section-title">
          <div class="icon-circle icon-deps">&#128279;</div>
          Partner Dependencies
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:13%">Work Stream</th>
                <th style="width:20%">Work Item</th>
                <th style="width:24%">Impact if Missed</th>
                <th style="width:25%">Mitigation</th>
                <th style="width:18%">Owner</th>
              </tr>
            </thead>
            <tbody>
{dep_table_rows(dependencies)}
            </tbody>
          </table>
        </div>
      </div>'''

html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{e(DASHBOARD_TITLE)} &mdash; {e(month_str)}</title>
<style>
  *, *::before, *::after {{ margin: 0; padding: 0; box-sizing: border-box; }}

  :root {{
    --navy: #002B5C;
    --navy-deep: #001a3a;
    --blue: #0078D4;
    --blue-light: #42A5F5;
    --red: #D13438;
    --orange: #FF8C00;
    --green: #107C10;
    --teal: #008272;
    --purple: #68217A;
    --white: #FFFFFF;
    --gray-50: #FAFBFC;
    --gray-100: #F3F4F6;
    --gray-200: #E8EBF0;
    --gray-300: #D1D5DB;
    --gray-400: #9CA3AF;
    --gray-600: #4B5563;
    --gray-800: #1F2937;
    --gray-900: #111827;
    --red-light: #FDE7E9;
    --green-light: #DDF6DF;
    --shadow-sm: 0 1px 2px rgba(0,0,0,0.04);
    --shadow-md: 0 4px 12px rgba(0,0,0,0.06);
    --shadow-lg: 0 8px 30px rgba(0,0,0,0.08);
    --radius-sm: 6px;
    --radius-md: 10px;
    --radius-lg: 16px;
    --radius-pill: 100px;
    --font: 'Segoe UI Variable', 'Segoe UI', -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
    --mono: 'Cascadia Code', 'Consolas', 'Courier New', monospace;
    --ease-out: cubic-bezier(0.16, 1, 0.3, 1);
  }}

  html {{ scroll-behavior: smooth; }}
  body {{
    font-family: var(--font);
    background: var(--gray-50);
    color: var(--gray-800);
    line-height: 1.5;
    -webkit-font-smoothing: antialiased;
    overflow-x: hidden;
  }}

  /* ==================== PROGRESS BAR ==================== */
  .sprint-progress {{ position: fixed; top: 0; left: 0; width: 100%; height: 3px; background: rgba(0,43,92,0.15); z-index: 1000; }}
  .sprint-progress-bar {{ height: 100%; width: {progress_pct}%; background: linear-gradient(90deg, var(--blue), var(--teal), var(--green)); border-radius: 0 3px 3px 0; }}

  /* ==================== HEADER ==================== */
  .header {{ background: var(--navy); position: relative; overflow: hidden; }}
  .header::before {{ content: ''; position: absolute; inset: 0; background: radial-gradient(ellipse 80% 50% at 70% 0%, rgba(0,120,212,0.15), transparent), radial-gradient(ellipse 60% 80% at 100% 100%, rgba(0,130,114,0.1), transparent); pointer-events: none; }}
  .header-mesh {{ position: absolute; inset: 0; opacity: 0.04; background-image: linear-gradient(rgba(255,255,255,0.3) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.3) 1px, transparent 1px); background-size: 48px 48px; mask-image: linear-gradient(135deg, black 30%, transparent 70%); -webkit-mask-image: linear-gradient(135deg, black 30%, transparent 70%); }}
  .header-inner {{ max-width: 1560px; margin: 0 auto; padding: 28px 48px 24px; position: relative; z-index: 1; display: flex; align-items: flex-start; justify-content: space-between; gap: 24px; }}
  .header-left {{ flex: 1; }}
  .header-label {{ display: inline-flex; align-items: center; gap: 6px; font-size: 11px; font-weight: 600; letter-spacing: 1.8px; text-transform: uppercase; color: rgba(255,255,255,0.45); margin-bottom: 8px; }}
  .header-label::before {{ content: ''; width: 16px; height: 2px; background: var(--blue); border-radius: 2px; }}
  .header-title {{ font-size: 32px; font-weight: 800; color: var(--white); letter-spacing: -0.5px; line-height: 1.15; }}
  .header-title span {{ background: linear-gradient(135deg, #60B0F4, #4ECDC4); -webkit-background-clip: text; -webkit-text-fill-color: transparent; background-clip: text; }}
  .header-date {{ font-size: 13px; color: rgba(255,255,255,0.4); margin-top: 6px; font-weight: 500; }}
  .header-right {{ display: flex; align-items: center; gap: 16px; padding-top: 8px; }}
  @keyframes dot-pulse {{ 0%, 100% {{ opacity: 1; transform: scale(1); }} 50% {{ opacity: 0.5; transform: scale(0.7); }} }}

  /* Health pill in tab button */
  .tab-health-pill {{ display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: var(--radius-pill); font-size: 9px; font-weight: 700; letter-spacing: 0.6px; text-transform: uppercase; color: var(--white); margin-right: 4px; vertical-align: middle; }}
  .tab-health-pill::before {{ content: ''; width: 5px; height: 5px; background: var(--white); border-radius: 50%; animation: dot-pulse 2s ease-in-out infinite; }}
  .tab-btn-health.active .tab-health-pill {{ box-shadow: 0 0 8px rgba(255,255,255,0.3); }}

  /* ==================== MAIN ==================== */
  .main {{ max-width: 1560px; margin: 0 auto; padding: 32px 48px 48px; }}

  /* ==================== KPI SECTION ==================== */
  .kpi-section {{ margin-bottom: 24px; }}
  .kpi-group {{ margin-bottom: 20px; }}
  .kpi-group-label {{ font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 1.5px; color: var(--gray-400); margin-bottom: 10px; padding-left: 2px; }}
  .kpi-groups {{ display: flex; gap: 32px; align-items: flex-start; }}
  .kpi-group {{ margin-bottom: 8px; }}
  .kpi-group-label {{ cursor: pointer; display: flex; align-items: center; gap: 8px; user-select: none; }}
  .kpi-group-label::after {{ content: ''; width: 16px; height: 16px; background: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='%239CA3AF' stroke-width='2.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E") no-repeat center; transition: transform 0.3s var(--ease-out); flex-shrink: 0; }}
  .kpi-group.collapsed .kpi-group-label::after {{ transform: rotate(-90deg); }}
  .kpi-group .kpi-row {{ transition: max-height 0.4s var(--ease-out), opacity 0.3s ease; max-height: 200px; opacity: 1; overflow: hidden; }}
  .kpi-group.collapsed .kpi-row {{ max-height: 0; opacity: 0; margin-top: 0; }}
  .kpi-group-label .kpi-group-summary {{ font-size: 10px; font-weight: 600; color: var(--gray-300); margin-left: 4px; display: none; }}
  .kpi-group.collapsed .kpi-group-label .kpi-group-summary {{ display: inline; }}
  .kpi-groups {{ display: flex; flex-direction: column; gap: 0; }}
  .kpi-group-features .kpi-row {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; }}
  .kpi-group-progress .kpi-row {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; }}
  .kpi-group-bugs .kpi-row {{ display: grid; grid-template-columns: repeat({6 if ICM_ENABLED else 4}, 1fr); gap: 14px; }}
  .kpi-card {{
    background: var(--white); border-radius: var(--radius-md); padding: 20px 16px 16px;
    position: relative; overflow: hidden; cursor: default;
    transition: all 0.35s var(--ease-out); box-shadow: var(--shadow-sm); border: 1px solid var(--gray-200);
  }}
  .kpi-card:hover {{ transform: translateY(-3px); box-shadow: var(--shadow-lg); border-color: transparent; }}
  .kpi-card::before {{ content: ''; position: absolute; top: 0; left: 0; right: 0; height: 3px; background: var(--card-accent); border-radius: 3px 3px 0 0; }}
  .kpi-card::after {{ content: ''; position: absolute; bottom: 0; right: 0; width: 64px; height: 64px; background: var(--card-accent); opacity: 0.04; border-radius: 50%; transform: translate(20%, 30%); transition: all 0.4s var(--ease-out); }}
  .kpi-card:hover::after {{ opacity: 0.08; transform: translate(10%, 20%) scale(1.3); }}
  .kpi-value {{ font-size: 36px; font-weight: 800; color: var(--card-accent); line-height: 1; letter-spacing: -1px; }}
  .kpi-link {{ text-decoration: none; color: inherit; display: block; }}
  .kpi-link:hover .kpi-value {{ text-decoration: underline; }}
  .kpi-label {{ font-size: 11px; font-weight: 600; color: var(--gray-400); text-transform: uppercase; letter-spacing: 0.8px; margin-top: 6px; }}

  /* ==================== RESOURCE UTILIZATION ==================== */
  .resource-row {{ display: grid; grid-template-columns: repeat({len(em_names) if em_names else 2}, 1fr); gap: 20px; }}
  .resource-card {{ background: var(--white); border-radius: var(--radius-lg); padding: 20px 24px; box-shadow: var(--shadow-sm); border: 1px solid var(--gray-200); transition: all 0.3s var(--ease-out); }}
  .resource-card:hover {{ box-shadow: var(--shadow-md); border-color: var(--gray-300); }}
  .resource-card-header {{ margin-bottom: 16px; }}
  .resource-card-name {{ font-size: 16px; font-weight: 700; color: var(--navy); }}
  .resource-card-kpis {{ display: flex; gap: 12px; align-items: center; }}
  .resource-kpi {{ text-align: center; flex: 1; padding: 10px 4px; background: var(--gray-50); border-radius: var(--radius-md); }}
  .resource-kpi-value {{ font-size: 28px; font-weight: 800; color: var(--rk-color); line-height: 1; letter-spacing: -0.5px; }}
  .resource-kpi-label {{ font-size: 9px; font-weight: 600; color: var(--gray-400); text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }}
  .resource-kpi-separator {{ width: 1px; height: 40px; background: var(--gray-200); flex-shrink: 0; }}

  /* ==================== TABS ==================== */
  .tabs-wrapper {{ margin-bottom: 0; position: sticky; top: 3px; z-index: 50; background: var(--gray-50); padding: 12px 0 0; }}
  .tabs-nav {{ display: inline-flex; background: var(--white); border-radius: var(--radius-pill); padding: 4px; box-shadow: var(--shadow-md); border: 1px solid var(--gray-200); position: relative; }}
  .tab-btn {{ position: relative; z-index: 2; padding: 10px 28px; border: none; background: transparent; font-family: var(--font); font-size: 13px; font-weight: 600; color: var(--gray-400); cursor: pointer; border-radius: var(--radius-pill); transition: color 0.3s ease; white-space: nowrap; }}
  .tab-btn:hover {{ color: var(--gray-800); }}
  .tab-btn.active {{ color: var(--white); }}
  .tab-slider {{ position: absolute; top: 4px; left: 4px; height: calc(100% - 8px); background: var(--navy); border-radius: var(--radius-pill); transition: all 0.4s var(--ease-out); z-index: 1; box-shadow: 0 2px 8px rgba(0,43,92,0.3); }}

  /* ==================== TAB CONTENT ==================== */
  .tab-content-area {{ margin-top: 24px; min-height: 300px; }}
  .tab-panel {{ display: none; animation: panel-in 0.45s var(--ease-out); }}
  .tab-panel.active {{ display: block; }}
  @keyframes panel-in {{ from {{ opacity: 0; transform: translateY(8px); }} to {{ opacity: 1; transform: translateY(0); }} }}

  /* ==================== TABLES ==================== */
  .table-container {{ background: var(--white); border-radius: var(--radius-lg); overflow: hidden; box-shadow: var(--shadow-md); border: 1px solid var(--gray-200); }}
  .table-header-bar {{ display: flex; align-items: center; justify-content: space-between; padding: 18px 24px; border-bottom: 1px solid var(--gray-200); background: var(--gray-50); }}
  .table-title {{ font-size: 15px; font-weight: 700; color: var(--navy); display: flex; align-items: center; gap: 8px; }}
  .table-title .count {{ background: var(--navy); color: var(--white); font-size: 11px; font-weight: 700; padding: 2px 10px; border-radius: var(--radius-pill); }}
  table {{ width: 100%; border-collapse: collapse; }}
  thead th {{ padding: 12px 16px; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: var(--gray-400); text-align: left; border-bottom: 2px solid var(--gray-200); background: var(--white); }}
  tbody tr {{ transition: all 0.2s ease; border-bottom: 1px solid var(--gray-100); }}
  tbody tr:last-child {{ border-bottom: none; }}
  tbody tr:hover {{ background: rgba(0,120,212,0.03); }}
  tbody td {{ padding: 14px 16px; font-size: 13px; color: var(--gray-800); vertical-align: middle; }}
  tbody td:first-child {{ font-weight: 600; }}
  .row-blocked {{ background: var(--red-light) !important; }}
  .row-blocked:hover {{ background: #fbd5d8 !important; }}
  .row-shipped {{ background: var(--green-light) !important; }}
  .row-shipped:hover {{ background: #c8eecb !important; }}

  /* Status Pills */
  .status-pill {{ display: inline-flex; align-items: center; gap: 5px; padding: 4px 12px; border-radius: var(--radius-pill); font-size: 11px; font-weight: 700; letter-spacing: 0.2px; white-space: nowrap; line-height: 1.4; }}
  .status-pill::before {{ content: ''; width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }}
  .status-not-started {{ background: #f0f0f0; color: #757575; }}
  .status-not-started::before {{ background: #BDBDBD; }}
  .status-pm-planning {{ background: #ECEFF1; color: #546E7A; }}
  .status-pm-planning::before {{ background: #90A4AE; }}
  .status-pm-spec {{ background: #ECEFF1; color: #546E7A; }}
  .status-pm-spec::before {{ background: #78909C; }}
  .status-in-design {{ background: #F3E5F5; color: var(--purple); }}
  .status-in-design::before {{ background: var(--purple); }}
  .status-design-complete {{ background: #F3E5F5; color: #7B1FA2; }}
  .status-design-complete::before {{ background: #7B1FA2; }}
  .status-in-engg {{ background: #E3F2FD; color: #1565C0; }}
  .status-in-engg::before {{ background: var(--blue-light); }}
  .status-engg-complete {{ background: #E8F5E9; color: #2E7D32; }}
  .status-engg-complete::before {{ background: #4CAF50; }}
  .status-engg-poc-complete {{ background: #E3F2FD; color: #1565C0; }}
  .status-engg-poc-complete::before {{ background: #42A5F5; }}
  .status-on-hold {{ background: #FFF3E0; color: #E65100; }}
  .status-on-hold::before {{ background: var(--orange); }}
  .status-blocked {{ background: var(--red-light); color: var(--red); }}
  .status-blocked::before {{ background: var(--red); }}
  .status-deprioritised {{ background: #f0f0f0; color: #757575; }}
  .status-deprioritised::before {{ background: #9E9E9E; }}
  .status-rolled-df {{ background: #E0F2F1; color: var(--teal); }}
  .status-rolled-df::before {{ background: var(--blue); }}
  .status-rolled-msit {{ background: #E0F2F1; color: var(--teal); }}
  .status-rolled-msit::before {{ background: var(--teal); }}
  .status-rolled-prod {{ background: #E8F5E9; color: var(--green); }}
  .status-rolled-prod::before {{ background: var(--green); }}

  .blocker-text {{ font-size: 12px; font-weight: 600; color: var(--red); display: flex; align-items: flex-start; gap: 5px; }}
  .blocker-text.at-risk {{ color: var(--orange); }}
  .blocker-icon {{ flex-shrink: 0; width: 14px; height: 14px; margin-top: 1px; }}

  /* ==================== BLOCKERS TAB ==================== */
  .sub-section {{ margin-bottom: 28px; }}
  .sub-section:last-child {{ margin-bottom: 0; }}
  .sub-section-title {{ font-size: 16px; font-weight: 700; color: var(--navy); margin-bottom: 14px; display: flex; align-items: center; gap: 10px; }}
  .sub-section-title .icon-circle {{ width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 15px; flex-shrink: 0; }}
  .icon-asks {{ background: #FFF3E0; }}
  .icon-deps {{ background: #E8EAF6; }}

  /* ==================== EPIC TRACKERS ==================== */
  .epic-section {{ margin-top: 48px; }}
  .epic-section-header {{ display: flex; align-items: center; gap: 12px; margin-bottom: 20px; }}
  .epic-section-title {{ font-size: 22px; font-weight: 800; color: var(--navy); letter-spacing: -0.3px; }}
  .epic-section-line {{ flex: 1; height: 2px; background: linear-gradient(90deg, var(--gray-200), transparent); border-radius: 2px; }}
  .epic-card {{ background: var(--white); border-radius: var(--radius-lg); overflow: hidden; box-shadow: var(--shadow-sm); border: 1px solid var(--gray-200); margin-bottom: 14px; transition: all 0.3s var(--ease-out); }}
  .epic-card:hover {{ box-shadow: var(--shadow-md); border-color: var(--gray-300); }}
  .epic-card-header {{ display: flex; align-items: center; justify-content: space-between; padding: 18px 24px; cursor: pointer; user-select: none; transition: background 0.2s ease; }}
  .epic-card-header:hover {{ background: var(--gray-50); }}
  .epic-card-left {{ display: flex; align-items: center; gap: 14px; }}
  .epic-icon {{ width: 40px; height: 40px; border-radius: var(--radius-md); display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0; }}
  .epic-icon-0 {{ background: linear-gradient(135deg, #E3F2FD, #BBDEFB); }}
  .epic-icon-1 {{ background: linear-gradient(135deg, #FFF3E0, #FFE0B2); }}
  .epic-icon-2 {{ background: linear-gradient(135deg, #E8F5E9, #C8E6C9); }}
  .epic-name {{ font-size: 15px; font-weight: 700; color: var(--gray-800); }}
  .epic-item-count {{ font-size: 12px; color: var(--gray-400); font-weight: 500; margin-top: 1px; }}
  .epic-card-right {{ display: flex; align-items: center; gap: 16px; }}
  .epic-mini-kpis {{ display: flex; gap: 6px; }}
  .mini-kpi {{ display: flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: var(--radius-pill); font-size: 11px; font-weight: 700; background: var(--gray-100); color: var(--gray-600); }}
  .mini-kpi-engg {{ background: #E3F2FD; color: #1565C0; }}
  .mini-kpi-rollout {{ background: #E0F2F1; color: var(--teal); }}
  .mini-kpi-shipped {{ background: #E8F5E9; color: var(--green); }}
  .mini-kpi-blocked {{ background: var(--red-light); color: var(--red); }}
  .chevron {{ width: 24px; height: 24px; transition: transform 0.35s var(--ease-out); color: var(--gray-400); }}
  .epic-card.open .chevron {{ transform: rotate(180deg); }}
  .epic-card-body {{ max-height: 0; overflow: hidden; transition: max-height 0.5s var(--ease-out); }}
  .epic-card.open .epic-card-body {{ max-height: 1200px; }}
  .epic-card-body-inner {{ padding: 0 24px 20px; border-top: 1px solid var(--gray-100); }}
  .epic-table {{ margin-top: 12px; }}
  .epic-table thead th {{ font-size: 10px; padding: 10px 12px; background: var(--gray-50); }}
  .epic-table tbody td {{ padding: 10px 12px; font-size: 12px; }}
  .priority-badge {{ display: inline-flex; align-items: center; justify-content: center; width: 28px; height: 20px; border-radius: 4px; font-size: 10px; font-weight: 800; font-family: var(--mono); }}
  .priority-P0 {{ background: var(--red); color: var(--white); }}
  .priority-P1 {{ background: var(--orange); color: var(--white); }}
  .priority-P2 {{ background: var(--blue); color: var(--white); }}
  .priority-P3 {{ background: var(--gray-300); color: var(--gray-600); }}
  .customer-pill {{ display: inline-flex; align-items: center; gap: 3px; padding: 2px 8px; border-radius: var(--radius-pill); font-size: 10px; font-weight: 700; background: #E8EAF6; color: var(--navy); letter-spacing: 0.2px; margin-left: 6px; vertical-align: middle; }}
  .customer-pill::before {{ content: ''; width: 5px; height: 5px; border-radius: 50%; background: var(--navy); }}

  /* Collapsible work item description */
  .wi-pitch {{ margin-top: 4px; cursor: pointer; }}
  .wi-pitch-toggle {{ font-size: 10px; color: var(--blue); font-weight: 600; letter-spacing: 0.2px; }}
  .wi-pitch-toggle:hover {{ text-decoration: underline; }}
  .wi-pitch-text {{ display: none; font-size: 11px; color: var(--gray-600); line-height: 1.5; margin-top: 4px; padding: 6px 8px; background: var(--gray-50); border-radius: var(--radius-sm); border-left: 2px solid var(--blue); font-weight: 400; }}
  .wi-pitch.open .wi-pitch-toggle {{ color: var(--gray-400); }}
  .wi-pitch.open .wi-pitch-text {{ display: block; }}

  /* ==================== MODAL ==================== */
  .modal-overlay {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.45); z-index: 999; align-items: center; justify-content: center; backdrop-filter: blur(3px); }}
  .modal-content {{ background: var(--white); border-radius: var(--radius-lg); width: 92%; max-width: 1300px; max-height: 85vh; overflow: hidden; box-shadow: 0 24px 80px rgba(0,0,0,0.2); display: flex; flex-direction: column; }}
  .modal-header {{ display: flex; align-items: center; justify-content: space-between; padding: 18px 28px; border-bottom: 1px solid var(--gray-200); background: var(--gray-50); }}
  .modal-title {{ font-size: 16px; font-weight: 700; color: var(--navy); display: flex; align-items: center; gap: 10px; }}
  .modal-title .count {{ background: var(--navy); color: var(--white); font-size: 11px; font-weight: 700; padding: 2px 10px; border-radius: var(--radius-pill); }}
  .modal-close {{ width: 32px; height: 32px; border: none; background: var(--gray-100); border-radius: 50%; cursor: pointer; font-size: 18px; color: var(--gray-600); display: flex; align-items: center; justify-content: center; transition: all 0.2s ease; }}
  .modal-close:hover {{ background: var(--gray-200); color: var(--gray-800); }}
  .modal-body {{ overflow-y: auto; padding: 0; }}

  /* ==================== FOOTER ==================== */
  .footer {{ max-width: 1560px; margin: 0 auto; padding: 24px 48px 40px; display: flex; align-items: center; justify-content: space-between; }}
  .footer-text {{ font-size: 11px; color: var(--gray-400); font-weight: 500; }}
  .footer-logo {{ display: flex; align-items: center; gap: 6px; font-size: 11px; font-weight: 700; color: var(--gray-300); letter-spacing: 0.5px; text-transform: uppercase; }}
  .footer-logo .dot {{ width: 4px; height: 4px; background: var(--blue); border-radius: 50%; }}

  /* ==================== RESPONSIVE ==================== */
  @media (max-width: 1400px) {{
    .kpi-groups {{ flex-direction: column; gap: 16px; }}
    .kpi-group-features .kpi-row {{ grid-template-columns: repeat(4, 1fr); }}
    .kpi-group-bugs .kpi-row {{ grid-template-columns: repeat({6 if ICM_ENABLED else 4}, 1fr); }}
    .resource-row {{ grid-template-columns: 1fr; }}
  }}
  @media (max-width: 900px) {{
    .header-inner {{ flex-direction: column; padding: 20px 24px; }}
    .main {{ padding: 20px 24px; }}
    .kpi-group-features .kpi-row {{ grid-template-columns: repeat(3, 1fr); }}
    .kpi-group-bugs .kpi-row {{ grid-template-columns: repeat({6 if ICM_ENABLED else 4}, 1fr); }}
    .tabs-nav {{ flex-wrap: wrap; border-radius: var(--radius-md); }}
    .tab-slider {{ border-radius: var(--radius-md); }}
    .epic-mini-kpis {{ display: none; }}
    .resource-card-kpis {{ flex-wrap: wrap; }}
  }}

  ::-webkit-scrollbar {{ width: 6px; }}
  ::-webkit-scrollbar-track {{ background: transparent; }}
  ::-webkit-scrollbar-thumb {{ background: var(--gray-300); border-radius: 3px; }}
  ::-webkit-scrollbar-thumb:hover {{ background: var(--gray-400); }}
</style>
</head>
<body>

<div class="sprint-progress">
  <div class="sprint-progress-bar" id="progressBar"></div>
</div>

<header class="header">
  <div class="header-mesh"></div>
  <div class="header-inner">
    <div class="header-left">
      <div class="header-label">{e(DASHBOARD_SUBTITLE)}</div>
      <h1 class="header-title">{e(DASHBOARD_TITLE)} &mdash; <span>{e(month_str)}</span></h1>
      <div class="header-date">Last updated: {e(today_str)} &nbsp;&bull;&nbsp; <strong style="color:rgba(255,255,255,0.7);">{days_left} days remaining in sprint</strong></div>
    </div>
    <div class="header-right">
    </div>
  </div>
</header>

<main class="main">

  <!-- KPI Cards — Visually Bucketed -->
  <div class="kpi-section">
    <div class="kpi-groups">
      <div class="kpi-group kpi-group-features" id="kpi-features" onclick="this.classList.toggle('collapsed')">
        <div class="kpi-group-label">Feature Items <span class="kpi-group-summary">({total} total, {in_rollout} rollout, {shipped} shipped, {blocked_count} blocked)</span></div>
        <div class="kpi-row">
          <div class="kpi-card" style="--card-accent: var(--blue); cursor:pointer;" onclick="event.stopPropagation();document.getElementById('modal-total').style.display='flex'">
            <div class="kpi-value" data-target="{total}">{total}</div>
            <div class="kpi-label">Total Items</div>
          </div>
          <div class="kpi-card" style="--card-accent: var(--teal); cursor:pointer;" onclick="event.stopPropagation();document.getElementById('modal-rollout').style.display='flex'">
            <div class="kpi-value" data-target="{in_rollout}">{in_rollout}</div>
            <div class="kpi-label">In Rollout (DF/MSIT)</div>
          </div>
          <div class="kpi-card" style="--card-accent: var(--green); cursor:pointer;" onclick="event.stopPropagation();document.getElementById('modal-shipped').style.display='flex'">
            <div class="kpi-value" data-target="{shipped}">{shipped}</div>
            <div class="kpi-label">Shipped (Preview/Prod)</div>
          </div>
          <div class="kpi-card" style="--card-accent: var(--red);">
            <div class="kpi-value" data-target="{blocked_count}">{blocked_count}</div>
            <div class="kpi-label">Blocked / At Risk</div>
          </div>
        </div>
      </div>
      <div class="kpi-group kpi-group-progress collapsed" id="kpi-progress" onclick="this.classList.toggle('collapsed')">
        <div class="kpi-group-label">Sprint Progress <span class="kpi-group-summary">({wins_count} met target, {moved_forward} moved forward, {stayed_same} no movement)</span></div>
        <div class="kpi-row">
          <div class="kpi-card" style="--card-accent: #107C10; cursor:pointer;" onclick="event.stopPropagation();document.getElementById('modal-wins').style.display='flex'">
            <div class="kpi-value" data-target="{wins_count}">{wins_count}</div>
            <div class="kpi-label">Met Target</div>
          </div>
          <div class="kpi-card" style="--card-accent: var(--teal);">
            <div class="kpi-value" data-target="{moved_forward}">{moved_forward}</div>
            <div class="kpi-label">Moved Forward</div>
          </div>
          <div class="kpi-card" style="--card-accent: var(--gray-400);">
            <div class="kpi-value" data-target="{stayed_same}">{stayed_same}</div>
            <div class="kpi-label">No Movement</div>
          </div>
        </div>
      </div>
{bug_kpi_group_html}
    </div>
  </div>

  <!-- Tabs -->
  <div class="tabs-wrapper">
    <nav class="tabs-nav" id="tabsNav">
      <div class="tab-slider" id="tabSlider"></div>
      <button class="tab-btn active tab-btn-health" data-tab="blocked-view" onclick="switchTab('blocked-view', this)">Sprint Health <span class="tab-health-pill" style="background:{health_color};">{e(sprint_health)}</span></button>
      <button class="tab-btn" data-tab="slipping" onclick="switchTab('slipping', this)">Slipping ETA ({len(slipping_items)})</button>
      <button class="tab-btn" data-tab="rollout-df" onclick="switchTab('rollout-df', this)">Rollout DF / MSIT ({len(rollout_df_msit)})</button>
      <button class="tab-btn" data-tab="rollout-prod" onclick="switchTab('rollout-prod', this)">Rollout Prod / Preview ({len(rollout_prod)})</button>
      <button class="tab-btn" data-tab="blockers" onclick="switchTab('blockers', this)">Blockers, Asks &amp; Dependencies</button>
{resource_tab_btn}
    </nav>
  </div>

  <div class="tab-content-area">

    <!-- Slipping ETA Tab -->
    <div class="tab-panel" id="panel-slipping">
      <div style="font-size:14px;font-weight:600;color:var(--gray-600);margin-bottom:16px;">{len(slipping_items)} items where Actual ETA differs from Committed ETA</div>
{grouped_section_html(slipping_items, 'slip') if slipping_items else '      <div style="padding:40px;text-align:center;color:var(--gray-400);font-size:15px;font-weight:500;">None &mdash; all items are on track with their committed ETA</div>'}
    </div>

    <div class="tab-panel" id="panel-rollout-df">
{grouped_section_html(rollout_df_msit, 'df')}
    </div>

    <div class="tab-panel" id="panel-rollout-prod">
{grouped_section_html(rollout_prod, 'prod')}
    </div>

    <div class="tab-panel" id="panel-blockers">
      <div class="sub-section">
        <div class="sub-section-title">
          <div class="icon-circle icon-asks">&#9889;</div>
          Top Asks / Decisions Needed
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:18%">Epic</th>
                <th style="width:22%">Work Item</th>
                <th style="width:42%">Ask / Decision Needed</th>
                <th style="width:18%">Owner</th>
              </tr>
            </thead>
            <tbody>
{asks_table_rows(asks)}
            </tbody>
          </table>
        </div>
      </div>
{deps_subsection_html}
    </div>

{"" if not EM_ENABLED else f"""    <!-- Resource Utilization Tab -->
    <div class="tab-panel" id="panel-resource">
      <div class="resource-row">
{resource_cards_html}
      </div>
    </div>"""}

    <!-- Blocked / At Risk Tab -->
    <div class="tab-panel active" id="panel-blocked-view">
      <div style="font-size:14px;font-weight:600;color:var(--gray-600);margin-bottom:16px;">{blocked_count} of {total} ({blocked_pct:.0f}%) items blocked or at risk</div>
{grouped_section_html(has_blocker, 'blk')}
    </div>

  </div>

  <!-- Epic Trackers -->
  <section class="epic-section">
    <div class="epic-section-header">
      <div class="epic-section-title">Epic Trackers</div>
      <div class="epic-section-line"></div>
    </div>
{epic_cards_html}
  </section>

</main>

{modals_html}

<footer class="footer">
  <div class="footer-text">Source: <a href="file:///{EXCEL_PATH.replace(chr(92), '/')}" style="color:var(--blue);text-decoration:none;font-weight:600;">{e(SOURCE_LABEL)}</a> &nbsp;|&nbsp; Generated {e(generated_str)}</div>
  <div class="footer-logo">
    <span class="dot"></span>
    {e(TEAM_NAME)}
    <span class="dot"></span>
  </div>
</footer>

<script>
  function switchTab(tabId, btn) {{
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.getElementById('panel-' + tabId).classList.add('active');
    btn.classList.add('active');
    updateSlider(btn);
  }}
  function updateSlider(btn) {{
    const slider = document.getElementById('tabSlider');
    const nav = document.getElementById('tabsNav');
    const navRect = nav.getBoundingClientRect();
    const btnRect = btn.getBoundingClientRect();
    slider.style.width = btnRect.width + 'px';
    slider.style.left = (btnRect.left - navRect.left) + 'px';
  }}
  function toggleEpic(id) {{
    document.getElementById(id).classList.toggle('open');
  }}
  function animateCounters() {{
    document.querySelectorAll('.kpi-value[data-target]').forEach(el => {{
      const target = parseInt(el.dataset.target);
      const duration = 1200;
      const start = performance.now();
      function tick(now) {{
        const elapsed = now - start;
        const progress = Math.min(elapsed / duration, 1);
        const eased = 1 - Math.pow(1 - progress, 4);
        el.textContent = Math.round(target * eased);
        if (progress < 1) requestAnimationFrame(tick);
      }}
      requestAnimationFrame(tick);
    }});
  }}
  window.addEventListener('DOMContentLoaded', () => {{
    const activeBtn = document.querySelector('.tab-btn.active');
    if (activeBtn) requestAnimationFrame(() => updateSlider(activeBtn));
    setTimeout(animateCounters, 300);
    window.addEventListener('resize', () => {{
      const active = document.querySelector('.tab-btn.active');
      if (active) updateSlider(active);
    }});
  }});
</script>
</body>
</html>'''

# ============================================================
# STEP 11: Write HTML file
# ============================================================
with open(HTML_OUTPUT_PATH, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\nDashboard saved to: {HTML_OUTPUT_PATH}")
print(f"\n=== Summary ===")
print(f"  Sprint: {sprint_tab_name}")
print(f"  Health: {sprint_health} ({blocked_pct:.0f}% blocked)")
print(f"  Items: {total} total, {in_engg} in engg, {in_rollout} in rollout, {shipped} shipped")
print(f"  Bugs: {p1_count} P1, {p2_count} P2, {resolved_count} resolved (14d)")
print(f"  Rollout DF/MSIT: {len(rollout_df_msit)} items")
print(f"  Rollout Prod: {len(rollout_prod)} items")
print(f"  Asks: {len(asks)}, Dependencies: {len(dependencies)}")
print(f"  Epics: {', '.join(f'{ep} ({len(epic_groups[ep])})' for ep in epic_order)}")
print(f"  EMs: {', '.join(em_names) if em_names else 'No EM tab found'}")
print(f"\nDone!")
