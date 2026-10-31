# ShiftSummary

Read-only dashboard of Skyline ER physician shift counts (shifts, overnights, weekends, early shifts) by scheduling year.

- `index.html` + `data.json`: the static site (GitHub Pages, main branch root).
- `build.py`: regenerates `data.json` from the source workbook `Shift Summary.xlsx` (not committed):
  `python3 build.py "Shift Summary.xlsx"` then copy `dist/data.json` here.

Names appear exactly as written in the spreadsheet (no merging).
