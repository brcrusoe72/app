# Setup

1. Install Python 3.10+.
2. Install dependencies:
   ```bash
   pip install openpyxl pytest
   ```
3. Optional (Windows Excel automation):
   ```bash
   pip install pywin32
   ```
4. Build or repair workbook:
   ```bash
   python scripts/build_or_repair_workbook.py
   ```

The script creates `excel/Shift_Flight_Deck.xlsm`, default rule snapshot, and core directory artifacts.
