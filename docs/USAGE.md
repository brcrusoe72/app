# Usage

## Excel flow

1. Run `python scripts/build_or_repair_workbook.py`.
2. Open `excel/Shift_Flight_Deck.xlsm` in Excel Desktop.
3. Use buttons/macros (when COM + VBA access is available): Validate, Analyze, Archive, Publish, Export Rules, Reload Rules.

## CLI flow

```bash
python scripts/analyze_workbook.py --workbook "excel/Shift_Flight_Deck.xlsm" --rules "data/rules.json"
python scripts/analyze_workbook.py --workbook "excel/Shift_Flight_Deck.xlsm" --export-rules
python scripts/archive_history.py --workbook "excel/Shift_Flight_Deck.xlsm"
python scripts/archive_history.py --workbook "excel/Shift_Flight_Deck.xlsm" --clear-current
python scripts/publish_reports.py --workbook "excel/Shift_Flight_Deck.xlsm"
```
