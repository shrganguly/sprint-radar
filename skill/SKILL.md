---
name: generate-shiproom-view
description: Generate a Shiproom Dashboard HTML from an Excel feature tracker and config file.
allowed-tools: Read, Bash, Glob, Grep
argument-hint: [path/to/config.json]
---

# Generate Shiproom View

Generate a styled HTML Shiproom Dashboard from an Excel feature tracker, ADO bug data, and ICM incidents.
Uses MSAL for Azure DevOps authentication (browser-based Microsoft login, no PAT needed).

## When invoked as a skill

1. **Locate config**: Use the argument as the config path. If no argument provided:
   - Search the current working directory for `config.json`
   - If not found, ask the user for the path
2. **Validate config**: Read the config JSON file and check:
   - `paths.excel_path` exists and points to a real file
   - If `ado.enabled` is true, check `ado.organization`, `ado.project`, `ado.area_path` are set
3. **Locate the script**: The `generate_html_dashboard.py` script lives in the same directory as the config file, or in the same directory as the Excel file. Search both.
4. **Run the generation**:
   ```bash
   python <script_dir>/generate_html_dashboard.py --config <config_path>
   ```
5. **Report**: Tell the user the output HTML path and open it

### Examples
- `/generate-shiproom-view config.json` -- use config.json in current directory
- `/generate-shiproom-view C:\Projects\RoB\config.json` -- absolute path
- `/generate-shiproom-view` -- auto-detect config.json in current directory

## Troubleshooting

- **"Tab not found"**: Check your `excel.sprint_tab_pattern` matches your tab naming convention
- **ADO auth popup**: First run opens browser for Microsoft login. Token is cached for subsequent runs.
- **No bug data**: Set `ado.enabled: false` to skip bugs entirely
- **ICM timeout**: ICM scraping requires an authenticated Edge session. Run `python _fetch_icm.py --config config.json` manually first to authenticate.
