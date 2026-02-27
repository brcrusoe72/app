# Shift Flight Deck Automation

Production-grade, automation-first Excel system for shift operations.

## Included deliverables

- `excel/Shift_Flight_Deck.xlsm` (generated locally; not committed)
- `excel/LineSight_Analyzer.xlam` (generated locally; not committed)
- `scripts/build_or_repair_workbook.py`
- `scripts/analyze_workbook.py`
- `scripts/archive_history.py`
- `scripts/publish_reports.py`
- `schemas/shift_flight_deck.schema.json`
- `data/history.sqlite` (created by archive script; not committed)
- `data/logs/` (runtime logs; not committed)
- `docs/SETUP.md`, `docs/USAGE.md`
- `tests/test_workbook_contract.py`

## How to run

```bash
python scripts/build_or_repair_workbook.py
python scripts/analyze_workbook.py --workbook "excel/Shift_Flight_Deck.xlsm" --rules "data/rules.json"
python scripts/analyze_workbook.py --workbook "excel/Shift_Flight_Deck.xlsm" --export-rules
python scripts/archive_history.py --workbook "excel/Shift_Flight_Deck.xlsm"
python scripts/publish_reports.py --workbook "excel/Shift_Flight_Deck.xlsm"
pytest -q
```

Then open workbook in Excel and use button macros on `Dash_Shift`.


## Binary artifact policy

Binary/generated artifacts are git-ignored to avoid push failures in environments that reject binary diffs. Generate workbook/data locally via the scripts.


## Quick walkthrough artifact

If you cannot run scripts locally, see `docs/SCRIPT_WALKTHROUGH_AND_EXAMPLE.md` for a captured run and sample report output.


## Merge conflict note

If a merge/rebase leaves conflict markers like `<<<<<<<`, `=======`, `>>>>>>>` in `README.md`, remove those markers and keep the final text block only.
A quick check is:

```bash
rg -n "^(<<<<<<<|=======|>>>>>>>)" README.md
```

If the command prints nothing, the conflict is resolved.
