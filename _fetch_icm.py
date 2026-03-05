"""Fetch ICM incident data using Playwright with Edge.
Supports optional --config argument to read ICM query URLs and output path from config.json.
Without --config, falls back to hardcoded defaults (backward compatible).
"""
import argparse, json, os, time
from datetime import datetime
from playwright.sync_api import sync_playwright

parser = argparse.ArgumentParser(description="Fetch ICM incidents via Playwright.")
parser.add_argument("--config", required=False, help="Path to config.json (optional)")
_args = parser.parse_args()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Load config if provided, otherwise use defaults
if _args.config:
    with open(_args.config, "r", encoding="utf-8") as f:
        _cfg = json.load(f)
    _icm = _cfg.get("icm", {})
    _paths = _cfg.get("paths", {})
    _active_url = _icm.get("active_query_url", "")
    _resolved_url = _icm.get("resolved_query_url", "")
    _data_file = _icm.get("data_file", "icm_data.json")
    _output_dir = _paths.get("html_output_dir") or os.path.dirname(os.path.abspath(_paths.get("excel_path", ""))) or SCRIPT_DIR
    OUTPUT_PATH = os.path.join(os.path.abspath(_output_dir), _data_file)
    print(f"Config loaded: {_args.config}")
else:
    _active_url = "https://portal.microsofticm.com/imp/v3/incidents/search/advanced?sl=agjfepzdktq"
    _resolved_url = "https://portal.microsofticm.com/imp/v3/incidents/search/advanced?sl=sphvcth050p"
    OUTPUT_PATH = os.path.join(SCRIPT_DIR, "icm_data.json")

QUERIES = []
if _active_url:
    QUERIES.append({"name": "active", "url": _active_url})
if _resolved_url:
    QUERIES.append({"name": "resolved", "url": _resolved_url})

if not QUERIES:
    print("No ICM query URLs configured. Nothing to fetch.")
    exit(0)

def extract_table(page, label):
    """Extract all rows from the ICM results table."""
    print(f"  Extracting table data for {label}...")

    # Get column headers
    headers = []
    header_cells = page.query_selector_all("th, [role='columnheader']")
    for h in header_cells:
        text = h.inner_text().strip()
        if text and text not in ("", "\n"):
            headers.append(text)
    print(f"  Headers: {headers}")

    # Get data rows
    incidents = []
    rows = page.query_selector_all("table tbody tr, tr[class*='row']")
    for row in rows:
        cells = row.query_selector_all("td, [role='gridcell']")
        if not cells or len(cells) < 3:
            continue
        texts = [c.inner_text().strip() for c in cells]
        # Build a dict mapping header to value
        if headers:
            # Skip checkbox column if present (first header might be empty)
            data_headers = [h for h in headers if h]
            offset = len(texts) - len(data_headers)
            if offset < 0:
                offset = 0
            incident = {}
            for i, h in enumerate(data_headers):
                idx = i + offset
                if idx < len(texts):
                    incident[h] = texts[idx]
            if incident:
                incidents.append(incident)
        else:
            incidents.append({"raw": texts})

    return incidents

with sync_playwright() as p:
    user_data_dir = os.path.join(os.environ["TEMP"], "playwright_edge_icm")
    context = p.chromium.launch_persistent_context(
        user_data_dir,
        channel="msedge",
        headless=False,
        viewport={"width": 1920, "height": 1080},
    )
    page = context.new_page()

    all_data = {}
    for query in QUERIES:
        label = query["name"]
        url = query["url"]
        print(f"\n=== Fetching {label} incidents ===")
        page.goto(url, wait_until="networkidle", timeout=90000)
        print(f"  Page title: {page.title()}")

        # If auth needed, wait for user
        if "identity" in page.title().lower() or "sign in" in page.title().lower():
            print("  *** Auth required - please sign in ***")
            page.wait_for_url("**/portal.microsofticm.com/imp/**", timeout=120000)
            print("  Auth complete, reloading query...")
            page.goto(url, wait_until="networkidle", timeout=90000)

        # Wait for results table or "items" count at bottom
        print("  Waiting for results to load...")
        try:
            page.wait_for_selector("text=/\\d+ of \\d+ items|\\d+ - \\d+ of \\d+ items|0 items/", timeout=30000)
        except:
            pass
        time.sleep(3)

        # Check if results loaded
        items_text = ""
        try:
            pager = page.query_selector("text=/items/")
            if pager:
                items_text = pager.inner_text()
                print(f"  Pager: {items_text}")
        except:
            pass

        # Screenshot
        screenshot_path = os.path.join(SCRIPT_DIR, f"icm_{label}.png")
        page.screenshot(path=screenshot_path, full_page=True)
        print(f"  Screenshot: icm_{label}.png")

        # Extract data
        incidents = extract_table(page, label)
        all_data[label] = incidents
        print(f"  Extracted: {len(incidents)} incidents")
        for inc in incidents[:5]:
            print(f"    {inc}")

    # Build summary and save
    output = {
        "exported_at": datetime.now().isoformat(),
        "active": all_data.get("active", []),
        "resolved": all_data.get("resolved", []),
        "summary": {
            "active_count": len(all_data.get("active", [])),
            "resolved_count": len(all_data.get("resolved", [])),
        }
    }

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print(f"\n=== Saved to {OUTPUT_PATH} ===")
    print(f"  Active: {output['summary']['active_count']}")
    print(f"  Resolved: {output['summary']['resolved_count']}")

    context.close()

print("\nDone!")
