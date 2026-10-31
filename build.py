#!/usr/bin/env python3
"""Extract Shift Summary.xlsx -> dist/data.json. Re-run whenever the spreadsheet changes:
    python3 build.py ["path/to/Shift Summary.xlsx"]
Reads cached formula values (data_only). Save the workbook in Excel before re-running."""
import sys, re, json, datetime, warnings
from pathlib import Path
import openpyxl
warnings.filterwarnings("ignore")
HERE = Path(__file__).parent
SRC = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "Shift Summary.xlsx"
OUT = HERE / "dist" / "data.json"  # site copy lives at repo root
MON = {"JAN":1,"FEB":2,"MAR":3,"APR":4,"MAY":5,"JUN":6,"JUL":7,"AUG":8,"SEP":9,"OCT":10,"NOV":11,"DEC":12}
SKIP_SHEETS = {"Night alc"}  # calendar-year night allocation worksheet, not a period summary

def mon(v):
    if not isinstance(v, str): return None
    return MON.get(v.strip().upper()[:3])

def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

def canon(name):
    # No merging/normalization: names are kept exactly as they appear in the Excel file.
    return str(name).strip()

def period_start(title, first_month):
    yrs = re.findall(r"\d{2,4}", title)
    y = int(yrs[0]); y = y + 2000 if y < 100 else y
    return y, first_month

def month_cols(row, lo=2, hi=19):
    return [(c.column, mon(c.value)) for c in row if lo <= c.column < hi and mon(c.value)]

def read_rows(ws, start):
    r = start
    while r <= ws.max_row:
        a = ws.cell(r, 1).value
        if a is None or (isinstance(a, str) and a.strip().upper() in ("TOTAL", "")):
            return
        yield r, a
        r += 1

def labels(cols, y0):
    out, y, prev = [], y0, None
    for _, m in cols:
        if prev is not None and m <= prev: y += 1
        out.append(f"{y}-{m:02d}"); prev = m
    return out

def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    periods, people = [], {}
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS: continue
        hdr = [r for r in ws.iter_rows() if len(month_cols(r)) >= 6]
        if len(hdr) < 3: continue
        on_h, we_h, sh_h = hdr[0], hdr[1], hdr[2]
        on_cols = month_cols(on_h, 2, 27)
        y0, m0 = period_start(ws.title, on_cols[0][1])
        months = labels(on_cols, y0)
        def block(h, cols, pair=False):
            res = {}
            for r, name in read_rows(ws, h[0].row + 1):
                vals = [num(ws.cell(r, c).value) for c, _ in cols]
                wk = [num(ws.cell(r, c + 1).value) for c, _ in cols] if pair else None
                res[canon(name)] = (str(name).strip(), vals, wk)
            return res
        on = block(on_h, on_cols, pair=True)
        we = block(we_h, month_cols(we_h, 2, 16))
        sh = block(sh_h, month_cols(sh_h, 2, 16))
        early_cols = [(c, m) for c, m in month_cols(sh_h, 20, 44)]
        ea = block(sh_h, early_cols) if early_cols else {}
        # early block may start later than period start (e.g. 2015-16 starts Nov)
        early_labels = {}
        if early_cols:
            first = early_cols[0][1]
            idx = next(i for i, l in enumerate(months) if int(l[5:]) == first)
            early_labels = dict(zip(range(len(early_cols)), months[idx:idx+len(early_cols)]))
        pid = ws.title
        roster = []
        for key in dict.fromkeys(list(sh) + list(on) + list(we)):
            raw = (sh.get(key) or on.get(key) or we.get(key))[0]
            def series(src, labs=None):
                if key not in src: return {}
                vals = src[key][1]
                labs = labs or months
                return {labs[i]: v for i, v in enumerate(vals) if i < len(labs) and v is not None}
            rec = {
                "shifts": series(sh), "nights": series(on), "weekends": series(we),
                "weekendNights": ({months[i]: v for i, v in enumerate(on[key][2]) if v is not None} if key in on else {}),
                "early": ({early_labels[i]: v for i, v in enumerate(ea[key][1]) if v is not None and i in early_labels} if key in ea else {}),
            }
            tot = {k: sum(v.values()) for k, v in rec.items()}
            tot["pctNights"] = round(tot["nights"]/tot["shifts"], 4) if tot["shifts"] else None
            tot["pctWeekends"] = round(tot["weekends"]/tot["shifts"], 4) if tot["shifts"] else None
            tot["pctEarly"] = round(tot["early"]/tot["shifts"], 4) if tot["shifts"] and rec["early"] else None
            if not any(tot[k] for k in ("shifts", "nights", "weekends")):  # roster placeholder with no work
                pass
            roster.append({"id": key, "label": raw, "monthly": rec, "totals": tot})
            people.setdefault(key, {"id": key, "name": key, "aliases": set()})["aliases"].add(raw)
        filled = [m for m in months if any(p["monthly"]["shifts"].get(m) for p in roster)]
        periods.append({"id": pid, "label": f"{months[0][:4]}–{months[-1][2:4]}", "sheet": ws.title,
                        "months": months, "monthsWithData": filled, "inProgress": len(filled) < len(months),
                        "physicians": roster})
    periods.sort(key=lambda p: p["months"][0], reverse=True)
    for p in people.values(): p["aliases"] = sorted(p["aliases"])
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps({"generated": datetime.datetime.now().isoformat(timespec="minutes"),
                               "source": SRC.name, "periods": periods,
                               "people": sorted(people.values(), key=lambda p: p["id"])}, indent=1))
    print(f"wrote {OUT}: {len(periods)} periods, {len(people)} people")
    for p in periods:
        print(" ", p["label"], p["sheet"], len(p["physicians"]), "rows; months w/ data", len(p["monthsWithData"]),
              "total shifts", sum(x["totals"]["shifts"] for x in p["physicians"]))

if __name__ == "__main__": main()
