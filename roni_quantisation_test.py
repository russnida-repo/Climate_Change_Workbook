#!/usr/bin/env python3
"""How much did the 1-decimal RONI series cost the +0.06 structural uncertainty?

HISTORICAL, and kept as the record of a prediction. This was run before the
ascii file was ingested, to decide whether re-ingesting was worth it. It
predicted a shift of about 0.0002, upward. The actual re-ingestion in v1.7.10
moved the figure 0.0001 downward: the magnitude was right and well inside this
test's stated standard deviation of 0.0004, the sign of the mean shift was not.

CPC publishes RONI at two decimals in RONI.v5.ascii.txt; the workbook carries the
one-decimal HTML table, so each value is the true value rounded. The true value is
therefore the stored value plus an unknown offset in roughly [-0.05, +0.05].

This draws that offset many times, writes a perturbed CSV each time, re-runs
open_items.py against it, and reports how far the RONI headline moves. It is an
upper bound on what re-ingesting the ascii file can change, computable without the
file itself. It does not replace the re-ingestion, which is still needed because
the workbook's stated premise is false and the CHECK should close against a
checksummed download.

Usage: python3 roni_quantisation_test.py [draws]
"""
import os, random, re, statistics, subprocess, sys, tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
# the SUPERSEDED one-decimal file: this test estimates what that precision cost,
# so it must read the rounded series, not the two-decimal one that replaced it
SRC = HERE / "roni_v5_cpc_1dp_superseded.csv"
WB = HERE / "Climate_Change_Analysis_2.xlsx"
SCRIPT = HERE / "open_items.py"
DRAWS = int(sys.argv[1]) if len(sys.argv) > 1 else 40
SEED = 20261023
ROW = re.compile(r"^\s*RONI v5\s+(-?\d+\.\d+)\s+(-?\d+\.\d+)")


def read_src():
    head, rows = [], []
    with open(SRC, encoding="utf-8-sig") as fh:
        for line in fh:
            if line.startswith("#") or line.startswith("year"):
                head.append(line.rstrip("\n"))
            elif line.strip():
                rows.append(line.rstrip("\n").split(","))
    return head, rows


def write_perturbed(path, head, rows, rng):
    with open(path, "w", newline="") as fh:
        for h in head:
            fh.write(h + "\n")
        for r in rows:
            out = [r[0]]
            for v in r[1:]:
                v = v.strip()
                if v == "":
                    out.append("")
                else:
                    out.append(f"{float(v) + rng.uniform(-0.05, 0.05):.2f}")
            fh.write(",".join(out) + "\n")


def run(csv_path):
    """Return (monthly degC per doubling, annual b, annual b x F2x).

    'RONI v5' labels a row in both the monthly and the annual block, so take both
    matches rather than the first: the first attempt at this read the monthly row
    and reported it as the annual headline.
    """
    p = subprocess.run([sys.executable, "-I", str(SCRIPT), str(WB), str(csv_path)],
                       capture_output=True, text=True, timeout=600)
    hits = [(float(m.group(1)), float(m.group(2)))
            for m in (ROW.match(l) for l in p.stdout.splitlines()) if m]
    assert len(hits) == 2, f"expected two RONI v5 rows, found {len(hits)}"
    monthly, annual = hits
    assert 2.0 < monthly[0] < 4.0, f"monthly row looks wrong: {monthly}"
    assert 0.3 < annual[0] < 0.6 and 1.3 < annual[1] < 2.0, f"annual row looks wrong: {annual}"
    return monthly[0], annual[0], annual[1]


head, rows = read_src()
print(f"source: {len(rows)} year rows, one decimal")

base_m, base_b, base_h = run(SRC)
print(f"baseline, as stored : monthly {base_m:.4f} degC/doubling | annual b {base_b:.4f} | annual b x F2x {base_h:.4f}")

rng = random.Random(SEED)
ms, bs, hs = [], [], []
tmpdir = tempfile.mkdtemp(prefix="roni_q_")
for i in range(DRAWS):
    p = os.path.join(tmpdir, f"roni_{i:03d}.csv")
    write_perturbed(p, head, rows, rng)
    m_, b, h = run(p)
    ms.append(m_); bs.append(b); hs.append(h)
    os.remove(p)
os.rmdir(tmpdir)

print(f"\n{DRAWS} draws of the lost precision, uniform in [-0.05, +0.05] per value:")
for name, base, xs in (("monthly model, degC per doubling", base_m, ms),
                       ("annual b, degC per W/m2", base_b, bs),
                       ("annual b x F2x, degC per doubling", base_h, hs)):
    mean = statistics.fmean(xs)
    sd = statistics.stdev(xs)
    print(f"  {name}")
    print(f"    stored {base:.4f} | perturbed mean {mean:.4f} | sd {sd:.4f} "
          f"| range {min(xs):.4f} to {max(xs):.4f}")
    print(f"    mean shift from the stored value: {mean - base:+.4f}")

onib = 1.6067
print(f"\nthe structural uncertainty this feeds, RONI minus ONI baseline {onib:.4f}:")
print(f"    as stored  : {base_h - onib:+.4f}")
print(f"    perturbed  : {statistics.fmean(hs) - onib:+.4f} "
      f"(sd {statistics.stdev(hs):.4f}, range {min(hs) - onib:+.4f} to {max(hs) - onib:+.4f})")
print("\nRead: the spread is what recovering the second decimal could move the figure by.")
