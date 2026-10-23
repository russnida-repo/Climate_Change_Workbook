#!/usr/bin/env python3
"""Verify the RONI re-ingestion independently of the writeup that proposed it.

Checks, all against the downloaded files rather than against any claim:
  1. the new two-decimal CSV reproduces RONI.v5.ascii.txt exactly
  2. the superseded one-decimal CSV is the rounding of those values, and which
     of five rounding conventions explains all 917
  3. RONI.v5.ascii.txt and RONI.ascii.txt are different series, quantified
"""
import decimal, statistics, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V5 = HERE / "RONI.v5.ascii.txt"             # the ERSSTv5 basis, the series used
OP = HERE / "RONI.ascii.txt"                # the operational series, for comparison
NEW = HERE / "roni_v5_cpc.csv"              # the current two-decimal ingestion
OLD = HERE / "roni_v5_cpc_1dp_superseded.csv"   # the 2026-10-22 HTML transcription
SEASONS = ["DJF", "JFM", "FMA", "MAM", "AMJ", "MJJ",
           "JJA", "JAS", "ASO", "SON", "OND", "NDJ"]


def read_ascii(path):
    out = {}
    with open(path, newline="") as fh:
        for i, line in enumerate(fh):
            p = line.split()
            if i == 0 or len(p) != 3:
                continue
            out[(int(p[1]), p[0])] = float(p[2])
    return out


def read_csv_grid(path):
    out = {}
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            if line.startswith("#"):
                continue
            p = [x.strip() for x in line.rstrip("\n").split(",")]
            if not p or p[0] in ("", "year"):
                continue
            yr = int(p[0])
            for s, v in zip(SEASONS, p[1:]):
                if v != "":
                    out[(yr, s)] = float(v)
    return out


v5, op = read_ascii(V5), read_ascii(OP)
new, old = read_csv_grid(NEW), read_csv_grid(OLD)
print(f"RONI.v5.ascii.txt  {len(v5)} values   RONI.ascii.txt  {len(op)} values")
print(f"new 2dp CSV        {len(new)} values   superseded 1dp CSV {len(old)} values")

# ---- 1. the new CSV must be the ascii file, exactly
assert set(new) == set(v5), (f"season sets differ: only in CSV {sorted(set(new)-set(v5))[:5]}, "
                             f"only in ascii {sorted(set(v5)-set(new))[:5]}")
bad = [(k, new[k], v5[k]) for k in v5 if abs(new[k] - v5[k]) > 1e-12]
print(f"\n1. new CSV against the ascii file: {len(v5)} seasons, mismatches {len(bad)}")
for k in bad[:5]:
    print("   ", k)

# ---- 2. which rounding convention explains the superseded transcription
def r_half_away(x):      # 0.5 -> 1, -0.5 -> -1
    return float(decimal.Decimal(repr(x)).quantize(decimal.Decimal("0.1"),
                 rounding=decimal.ROUND_HALF_UP))
def r_half_toward(x):    # 0.5 -> 0, -0.5 -> 0
    return float(decimal.Decimal(repr(x)).quantize(decimal.Decimal("0.1"),
                 rounding=decimal.ROUND_HALF_DOWN))
def r_half_even(x):
    return float(decimal.Decimal(repr(x)).quantize(decimal.Decimal("0.1"),
                 rounding=decimal.ROUND_HALF_EVEN))
def _directed(x, tie_up):
    """Round to one decimal, breaking exact ties along the number line.

    decimal has no ROUND_HALF_CEILING/FLOOR, so detect the tie and step it.
    """
    d = decimal.Decimal(repr(x)) * 10
    fl = d.to_integral_value(rounding=decimal.ROUND_FLOOR)
    if d - fl == decimal.Decimal("0.5"):
        n = fl + 1 if tie_up else fl
    else:
        n = d.to_integral_value(rounding=decimal.ROUND_HALF_EVEN)
    return float(n / 10)

def r_half_pinf(x):      # ties up the number line: -0.45 -> -0.4
    return _directed(x, True)
def r_half_ninf(x):      # ties down the number line: 0.45 -> 0.4
    return _directed(x, False)

common = sorted(set(old) & set(v5))
print(f"\n2. superseded 1dp CSV, {len(common)} seasons in common with the ascii file")
for name, fn in (("half away from zero", r_half_away), ("half toward zero", r_half_toward),
                 ("half to even", r_half_even), ("half toward +infinity", r_half_pinf),
                 ("half toward -infinity", r_half_ninf)):
    ok = sum(1 for k in common if abs(fn(v5[k]) - old[k]) < 1e-9)
    flag = "  <-- explains all" if ok == len(common) else ""
    print(f"   {name:<24} {ok} of {len(common)}{flag}")
maxdiff = max(abs(old[k] - v5[k]) for k in common)
ties = sum(1 for k in common if abs(abs(v5[k] * 100) % 10 - 5) < 1e-9)
print(f"   max |1dp - 2dp| = {maxdiff:.4f} (quantisation bound 0.05)")
print(f"   values sitting exactly on a .x5 tie: {ties}")

# ---- 3. are the two published files the same series?
both = sorted(set(v5) & set(op))
diffs = [v5[k] - op[k] for k in both]
ident = sum(1 for d in diffs if abs(d) < 1e-12)
flip = sum(1 for k in both
           if (v5[k] >= 0.5) != (op[k] >= 0.5) or (v5[k] <= -0.5) != (op[k] <= -0.5))
worst = max(both, key=lambda k: abs(v5[k] - op[k]))
print(f"\n3. the two published files over {len(both)} shared seasons")
print(f"   identical values      : {ident} ({100*ident/len(both):.1f}%)")
print(f"   mean |difference|     : {statistics.fmean(abs(d) for d in diffs):.4f}")
print(f"   mean signed difference: {statistics.fmean(diffs):+.4f}")
print(f"   max  |difference|     : {abs(v5[worst]-op[worst]):.2f} at {worst[1]} {worst[0]}"
      f"  (v5 {v5[worst]:+.2f}, operational {op[worst]:+.2f})")
print(f"   seasons on opposite sides of the +/-0.5 threshold: {flip} "
      f"({100*flip/len(both):.1f}%)")
print(f"   correlation           : "
      f"{statistics.correlation([v5[k] for k in both],[op[k] for k in both]):.4f}")
_ord = {s: i for i, s in enumerate(SEASONS)}
_last = lambda d: max(d, key=lambda k: (k[0], _ord[k[1]]))
print(f"   coverage: v5 ends {_last(v5)[1]} {_last(v5)[0]}, "
      f"operational ends {_last(op)[1]} {_last(op)[0]}; v5 lags by {len(op) - len(v5)} seasons")
top = sorted(both, key=lambda k: -abs(v5[k]-op[k]))[:5]
print("   largest divergences:", ", ".join(f"{k[1]} {k[0]} ({v5[k]-op[k]:+.2f})" for k in top))
