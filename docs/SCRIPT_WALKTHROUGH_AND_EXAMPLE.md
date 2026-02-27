# Shift Flight Deck Scripts: What They Do + Concrete Example

This document is a **readable artifact** for environments where you cannot run scripts locally.

## 1) What each script does

### `scripts/build_or_repair_workbook.py`
Creates or repairs `excel/Shift_Flight_Deck.xlsm` with the required workbook contract:
- Creates required sheets (`Parameters`, `Schedule_Entry`, `Hourly_Log`, `Downtime_Log`, `Dash_Shift`, `Dash_Trends`, `Profiles`, `Analysis_Report`, `Rules_Authoring`).
- Ensures required Excel tables exist (`tblLines`, `tblStandards`, `tblMachines`, `tblOperators`, `tblSchedule`, `tblHourly`, `tblDowntime`, `tblRules`).
- Seeds default sample data when workbook is empty.
- Applies rule authoring dropdown validations (Enabled/Severity/Scope).
- Builds basic dashboard/report placeholders.
- Exports default rules snapshot to `data/rules.json`.
- On Windows with COM available, it attempts VBA injection for button macros.

### `scripts/analyze_workbook.py`
Runs deterministic analysis and rules evaluation:
- Reads workbook tables.
- Lints rule rows in `tblRules` (required fields/enums/DSL parse).
- Evaluates deterministic DSL conditions (no ML).
- Writes an `Analysis_Report` sheet with sections:
  - Data Quality
  - Schedule Integrity
  - Standards Coverage
  - Operational Risks
  - Recommended Actions (ranked)
  - Rules Engine Coaching Prompts
  - Rule Lint
- Can export rules with:
  - `--export-rules` -> writes `data/rules.json`
  - logs to `data/logs/rules_export.log`.

### `scripts/archive_history.py`
Archives the current workbook logs into `data/history.sqlite`:
- Upserts schedule/hourly/downtime into `schedule_log`, `hourly_log`, `downtime_log`.
- Dedupe key is `RowID`.
- Optional `--clear-current` removes active rows after archive.

### `scripts/publish_reports.py`
Publishes shift artifacts:
- Saves timestamped workbook copy to `exports/`.
- Attempts PDF export of `Dash_Shift` and `Dash_Trends` via COM (Windows Excel).
- Writes `Shift_Summary_<timestamp>.txt` in `exports/`.
- Logs action to `data/logs/publish.log`.

---

## 2) Example run (captured in this repo)

### Build command + output
```bash
python scripts/build_or_repair_workbook.py
```

Captured output:

```text
Workbook ready: /workspace/app/excel/Shift_Flight_Deck.xlsm
win32com unavailable; VBA/button injection skipped
```

### Analyze command + output
```bash
python scripts/analyze_workbook.py --workbook excel/Shift_Flight_Deck.xlsm --rules data/rules.json
```

Captured output:

```text
Analyze complete
```

### Archive command + output
```bash
python scripts/archive_history.py --workbook excel/Shift_Flight_Deck.xlsm
```

Captured output:

```text
Archive complete
```

### Publish command + output
```bash
python scripts/publish_reports.py --workbook excel/Shift_Flight_Deck.xlsm
```

Captured output:

```text
Publish complete
```

---

## 3) Example `Analysis_Report` rows (captured)

```text
1 ['Analysis Report', None, None]
3 ['Data Quality', None, None]
4 ['- Hourly rows: 3', None, None]
5 ['- Downtime rows: 1', None, None]
6 ['- Rules source: workbook', None, None]
8 ['Schedule Integrity', None, None]
9 ['- Hourly rows without schedule: 0', None, None]
11 ['Standards Coverage', None, None]
12 ['- Rows missing standards: 3', None, None]
14 ['Operational Risks', None, None]
15 ['- Triggered prompts: 1', None, None]
17 ['Recommended Actions (ranked)', None, None]
18 ['- Urgent: Add standard immediately or switch to approved alternate SKU standard. (Line 1,SKU-001)', None, None]
20 ['Rules Engine Coaching Prompts', None, None]
21 ['RuleID', 'Severity', 'Trigger']
22 ['R2_MISSING_STANDARD', 'Urgent', 'Standards missing for active run']
24 ['Rule Lint', None, None]
25 ['No linter issues', None, None]
```

---

## 4) If you only need the shortest path

1. `python scripts/build_or_repair_workbook.py`
2. `python scripts/analyze_workbook.py --workbook excel/Shift_Flight_Deck.xlsm --rules data/rules.json`
3. `python scripts/archive_history.py --workbook excel/Shift_Flight_Deck.xlsm`
4. `python scripts/publish_reports.py --workbook excel/Shift_Flight_Deck.xlsm`

