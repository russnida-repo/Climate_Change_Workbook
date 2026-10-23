#!/usr/bin/env python3
"""Prove which season each row of Open Items column C holds.

Column C carries the RONI series as pasted literals, so nothing in the sheet
states the convention. This derives it from the month labels and then proves it,
by requiring that all 819 rows reproduce the shipped CSV exactly. It was written
to run before v1.7.10 rewrote that column, against the then-current one-decimal
file; it now checks against the two-decimal file the workbook holds, and
separately confirms that the superseded transcription is the one-decimal rounding
of the same seasons.

A standing check, not a one-shot: if column C and the CSV ever drift apart, this
fails rather than writing a row map.
"""
import json, sys
import openpyxl
from pathlib import Path

SEASONS = ["DJF", "JFM", "FMA", "MAM", "AMJ", "MJJ",
           "JJA", "JAS", "ASO", "SON", "OND", "NDJ"]
HERE = Path(__file__).resolve().parent
WB = HERE / "Climate_Change_Analysis_2.xlsx"
OLD = HERE / "roni_v5_cpc_1dp_superseded.csv"
NEW = HERE / "roni_v5_cpc.csv"


def read_grid(path):
    out = {}
    for line in open(path, encoding="utf-8-sig"):
        if line.startswith("#"):
            continue
        p = [x.strip() for x in line.rstrip("\n").split(",")]
        if not p or p[0] in ("", "year"):
            continue
        for s, v in zip(SEASONS, p[1:]):
            if v != "":
                out[(int(p[0]), s)] = float(v)
    return out


cur, sup = read_grid(NEW), read_grid(OLD)
f = openpyxl.load_workbook(WB)
v = openpyxl.load_workbook(WB, data_only=True)
oi = f["Open Items"]
vol = v["Volcanic"]

rows = [r for r in range(2, 900)
        if isinstance(oi[f"C{r}"].value, (int, float))]
print(f"numeric literals in column C: {len(rows)} rows, {rows[0]} to {rows[-1]}")

# column A mirrors Volcanic!A, which carries the month label
sample = [(r, vol[f"A{r}"].value) for r in (rows[0], rows[0] + 1, rows[-1])]
print("Volcanic!A at those rows:", sample)

# the month index for a row, from the Volcanic date column
def month_of(r):
    """Volcanic!A holds the month as the string 'M/YYYY', not a date."""
    d = vol[f"A{r}"].value
    if hasattr(d, "year"):
        return d.year, d.month
    if isinstance(d, str) and "/" in d:
        mo, yr = d.split("/")
        return int(yr), int(mo)
    return None


print("parsed as dates:", [month_of(r) for r in (rows[0], rows[0] + 1, rows[-1])])

mapping, misses = {}, []
for r in rows:
    ym = month_of(r)
    if ym is None:
        misses.append((r, "unparseable date"))
        continue
    yr, mo = ym
    key = (yr, SEASONS[mo - 1])          # season centred on this month
    if key not in cur:
        misses.append((r, f"no CSV value for {key}"))
        continue
    stored = round(float(oi[f"C{r}"].value), 10)
    if abs(stored - cur[key]) > 1e-9:
        misses.append((r, f"{key}: column C {stored} vs CSV {cur[key]}"))
        continue
    mapping[r] = key

print(f"\nrows whose stored value matches the shipped CSV under the centre-month "
      f"convention: {len(mapping)} of {len(rows)}")
if misses:
    print("misses:")
    for m in misses[:12]:
        print("   ", m)

if len(mapping) == len(rows):
    out = {str(r): [mapping[r][0], mapping[r][1], cur[mapping[r]]] for r in mapping}
    missing2 = [r for r in mapping if mapping[r] not in cur]
    assert not missing2, f"the shipped CSV lacks {len(missing2)} of the mapped seasons"
    json.dump(out, open(HERE / "roni_row_map.json", "w"), indent=0)
    print(f"\nmapping proven on all {len(rows)} rows and written to roni_row_map.json")
    import decimal

    def r1(x, tie_up=True):
        d = decimal.Decimal(repr(x)) * 10
        fl = d.to_integral_value(rounding=decimal.ROUND_FLOOR)
        if d - fl == decimal.Decimal("0.5"):
            n = fl + 1 if tie_up else fl
        else:
            n = d.to_integral_value(rounding=decimal.ROUND_HALF_EVEN)
        return float(n / 10)

    covered = [r for r in mapping if mapping[r] in sup]
    ok = sum(1 for r in covered if abs(r1(cur[mapping[r]]) - sup[mapping[r]]) < 1e-9)
    print(f"superseded one-decimal file: {len(covered)} of these seasons present, "
          f"{ok} of them the exact rounding (ties toward +infinity) of the shipped value")
    assert ok == len(covered), "the superseded file is not the rounding of the shipped one"
else:
    print("\nMAPPING NOT PROVEN: refusing to write the row map")
    sys.exit(1)
