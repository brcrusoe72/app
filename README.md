# Production Schedule Consolidator

Converts a Daily Production Schedule workbook (one worksheet per date) into a single workbook organized by production line.

## Input

- **Workbook**: `Daily Production Schedule 8.8.25.xlsx` (197 date sheets)
- Each sheet contains sections for Lines 1–5 with SKU rows including cases, shifts, completion data, and notes

## Output

- **Workbook**: `Production_Schedule_By_Line.xlsx`

| Sheet | Contents |
|---|---|
| README | Processing description and statistics |
| Assumptions & Data Issues | Issue log with severity, sheet reference, and action taken |
| Summary | Date × Line rollup with weighted avg % complete |
| Line 1 – Line 5 | Excel Tables with all schedule rows for that line |

### Line sheet columns

Date, SourceSheet, Line, SKU, SKU_RawText, Description, Cases_Planned, Shifts_Planned, Target_Per_Shift, Cases_Completed, Percent_Complete, Notes, WorkOrderMade, ExtraFields_JSON

## Usage

```bash
pip install openpyxl pandas
python consolidate_schedules.py
```

Place the input workbook at `/mnt/data/Daily Production Schedule 8.8.25.xlsx`. The script writes the output to `/mnt/data/Production_Schedule_By_Line.xlsx`.

## Processing stats

| Metric | Value |
|---|---|
| Sheets processed | 197 |
| Total rows extracted | 3,340 |
| Line 1 | 669 |
| Line 2 | 510 |
| Line 3 | 1,019 |
| Line 4 | 706 |
| Line 5 | 436 |

## Key behaviors

- **Date detection**: Prefers real datetime cells in header rows; falls back to parsing sheet names (M.D.YY and MM.DD.YYYY formats)
- **SKU parsing**: Extracts the first 6+ digit number from column B; short (4-digit) SKUs accepted with a logged note
- **Description**: Uses ` / ` as an explicit separator when present; otherwise takes everything after the SKU token
- **Percent Complete**: Always recomputed as `Cases_Completed / Cases_Planned` (ignores original `#DIV/0!` formulas)
- **Duplicates**: Flagged by (Date, Line, SKU, Cases_Planned, Shifts_Planned) but kept in output
- **Conditional formatting**: <80% complete (yellow), missing target per shift (red), missing SKU (orange)
