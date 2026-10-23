#!/usr/bin/env python3
"""
erf_uncertainty.py

Propagates IGCC 2025 effective-radiative-forcing uncertainty through the annual
anthropogenic-ERF regression of Climate_Change_Analysis_2.xlsx, producing the
figures pasted into Sensitivity!AB37:AD47.

The regression is the one at Forcing!T13:T17:

    mean annual anomaly = a + b * (anthropogenic ERF) + c * (ONI lag 3),
    1959-2025, 67 years

and the reported quantity is b * F2x, degC per CO2 doubling, with F2x = 3.93
W/m2 (AR6 best estimate, 5-95% half-width 0.47).

PROVENANCE, READ THIS FIRST
---------------------------
This file is a RECONSTRUCTION, written 2026-10-22 from the method documented at
Sensitivity!AA35:AA53 and the dataset entries at Sources rows 9 and 10. The
original script of this name was run in September 2026 and was not kept with the
workbook. It reproduces the pasted figures as follows:

  Exact, to the precision the workbook stores (every deterministic case except the
  first, which is a live link rather than a pasted value):
    Sensitivity!AB37:AC37   best estimate, 0.4088 / 1.607. These two cells are
                            =Forcing!T14 and =Forcing!T17, not pasted, and the
                            workbook regresses on ERF values rounded to 4 decimals
                            on the Forcing tab, so this script agrees to about
                            2e-6 rather than exactly
    Sensitivity!AB38:AC39   anthropogenic total at its pointwise 5th and 95th
                            percentile
    Sensitivity!AB40:AC41   aerosol series scaled by 1.67 and 0.41
    Sensitivity!AB42:AC42   WMGHG series scaled by 0.91 and 1.09
    Sensitivity!AC43        the single-component b*F2x ranges
    Sensitivity!AC47:AD47   the quadrature row and its 90% and 95% intervals

  The Monte Carlo rows, at the full 20,000 draws, reproduce every published
  percentile exactly as printed, on this script's own seed rather than the
  original run's, with the standard deviations agreeing to 0.001:
    Sensitivity!AC44:AD44   1.46 / 1.60 / 1.77, sd 0.095 (this run 0.095)
    Sensitivity!AC45:AD45   1.50 / 1.60 / 1.72, sd 0.068 (this run 0.067)
    Sensitivity!AC46:AD46   1.43 / 1.60 / 1.78, sd 0.107 (this run 0.108),
                            2.5-97.5% 1.39-1.81 (this run 1.39-1.82)
  At a few thousand draws the fifth percentiles wander by about 0.01, so use the
  default 20,000 when checking against the workbook.

Input data
----------
From https://github.com/ClimateIndicator/forcing-timeseries (Zenodo
doi:10.5281/zenodo.10706344), the files named in Sources rows 9 and 10:

    output/ERF_best_aggregates.csv   best estimates, aggregated to the component
                                     level the workbook uses (incl. halogen,
                                     aerosol, anthro), 1750-2025
    output/ERF_p05_aggregates.csv    5th percentile, same components
    output/ERF_p95_aggregates.csv    95th percentile, same components

Pass the directory holding them as the second argument. The repository copy checked
on 2026-10-22 (commit 852f4f7) reproduces the ERF values pasted on the Forcing
tab to 5e-5, which is that tab's 4-decimal rounding, so the vintage matches the
2026-09-21 retrieval the workbook records.

The anomaly and ONI series are read from the workbook itself (Forcing!B and
Forcing!D), since they come from GISTEMP and NOAA CPC by way of the other tabs.

Method
------
Eleven anthropogenic components are treated as independent: CO2, CH4, N2O,
halogen, aerosol-radiation_interactions, aerosol-cloud_interactions, O3,
contrails, land_use, BC_on_snow, H2O_stratospheric. For each, the 2019 relative
half-width is

    rel_sd_i = ((p95_i(2019) - p05_i(2019)) / 2 / 1.645) / |best_i(2019)|

and a draw multiplies that component's whole series by a factor ~ N(1, rel_sd_i).
The scaled components are summed to an anthropogenic total, the regression is
refitted, and b * F2x recorded.

Three Monte Carlo variants, as on the Sensitivity tab:
  (a) all 11 components sampled, F2x fixed at 3.93;
  (b) the same with CO2 held at its best estimate, isolating non-CO2 ERF
      uncertainty;
  (c) all 11 sampled, F2x scaled by the same draw as CO2 (they are the same
      physical quantity, so scaling CO2 down makes b larger and F2x smaller in
      nearly equal proportion), plus AR(1)-inflated regression noise on b.

Caveats carried over from the workbook (Sensitivity!AA52): independent scale
factors on best-estimate series ignore the ARI/ACI correlation and all shape
uncertainty, and the pointwise percentile rows are envelopes, not ensemble
members. The 1000-member IGCC ensemble is not published in the repository.

Usage
-----
  python3 erf_uncertainty.py <workbook.xlsx> <erf_csv_directory> [draws] [seed]

Dependencies: numpy, openpyxl.
"""

import csv
import os
import sys

import numpy as np
import openpyxl

F2X = 3.93
F2X_HALFWIDTH_5_95 = 0.47
Z90 = 1.645
DRAWS = 20_000
SEED = 20261022

COMPONENTS = ["CO2", "CH4", "N2O", "halogen",
              "aerosol-radiation_interactions", "aerosol-cloud_interactions",
              "O3", "contrails", "land_use", "BC_on_snow", "H2O_stratospheric"]
WMGHG = ["CO2", "CH4", "N2O", "halogen"]
AEROSOL = ["aerosol-radiation_interactions", "aerosol-cloud_interactions"]
YEAR0, YEAR1 = 1959, 2025
REF_YEAR = 2019


def read_erf(path):
    """{year: {component: value}} from one IGCC output CSV."""
    with open(path, newline="") as fh:
        rd = csv.reader(fh)
        hdr = next(rd)
        out = {}
        for row in rd:
            if not row or not row[0]:
                continue
            yr = int(float(row[0]))
            out[yr] = {hdr[i]: (float(row[i]) if row[i] not in ("", "nan") else np.nan)
                       for i in range(1, len(hdr))}
    return out


def slope(anom, erf, oni):
    """b, the coefficient on anthropogenic ERF."""
    A = np.column_stack([np.ones(len(erf)), erf, oni])
    beta, *_ = np.linalg.lstsq(A, anom, rcond=None)
    return beta[1]


def ck(label, got, want, tol, fmt="{:.4f}"):
    mark = "  ok" if want is None or abs(got - want) <= tol else "  DIFFERS"
    w = "" if want is None else f"   workbook {fmt.format(want)}"
    print(f"   {label:<50} {fmt.format(got)}{w}{mark}")


def main(wb_path, erf_dir, draws=DRAWS, seed=SEED):
    agg_b = read_erf(os.path.join(erf_dir, "ERF_best_aggregates.csv"))
    agg_l = read_erf(os.path.join(erf_dir, "ERF_p05_aggregates.csv"))
    agg_h = read_erf(os.path.join(erf_dir, "ERF_p95_aggregates.csv"))

    wb = openpyxl.load_workbook(wb_path, data_only=True)
    F = wb["Forcing"]
    S = wb["Sensitivity"]
    years, anom, oni = [], [], []
    for r in range(2, 69):
        y = F.cell(r, 1).value
        if isinstance(y, (int, float)):
            years.append(int(y))
            anom.append(float(F.cell(r, 2).value))
            oni.append(float(F.cell(r, 4).value))
    anom = np.array(anom)
    oni = np.array(oni)
    assert years[0] == YEAR0 and years[-1] == YEAR1, years[:1] + years[-1:]

    # per-component best-estimate series over the fit years
    series = {c: np.array([agg_b[y][c] for y in years]) for c in COMPONENTS}
    anthro_best = sum(series.values())
    anthro_wb = np.array([float(F.cell(r, 12).value) for r in range(2, 69)])
    print(f"Workbook: {wb_path}")
    print(f"ERF data: {erf_dir}")
    print(f"Fit years {years[0]}-{years[-1]} ({len(years)} years), "
          f"{len(COMPONENTS)} components\n")
    print("0. Inputs")
    ck("sum of 11 components vs Forcing!L, max abs diff",
       float(np.abs(anthro_best - anthro_wb).max()), 0.0, 2e-4, "{:.2e}")

    def cell(ref):
        v = S[ref].value
        return v if isinstance(v, (int, float)) else None

    b_best = slope(anom, anthro_best, oni)
    print("\n1. Best estimate   [Sensitivity!AB37:AC37]")
    print("   (the workbook regresses on the ERF values pasted on the Forcing tab, rounded")
    print("    to 4 decimals, so b agrees to about 2e-6 rather than exactly)")
    ck("b, degC per W/m2", b_best, cell("AB37"), 5e-4)
    ck("b * F2x, degC per doubling", b_best * F2X, cell("AC37"), 2e-3)

    print("\n2. Pointwise percentile envelopes   [Sensitivity!AB38:AC39]")
    for lab, agg, refs in (("5th", agg_l, ("AB38", "AC38")),
                           ("95th", agg_h, ("AB39", "AC39"))):
        erf = np.array([agg[y]["anthro"] for y in years])
        b = slope(anom, erf, oni)
        ck(f"anthropogenic total at its {lab} percentile: b", b, cell(refs[0]), 5e-4)
        ck(f"   b * F2x", b * F2X, cell(refs[1]), 5e-4, "{:.3f}")

    print("\n3. Scaled-series cases   [Sensitivity!AB40:AC42]")
    aer_ratio_lo = agg_l[REF_YEAR]["aerosol"] / agg_b[REF_YEAR]["aerosol"]
    aer_ratio_hi = agg_h[REF_YEAR]["aerosol"] / agg_b[REF_YEAR]["aerosol"]
    wm_lo = agg_l[REF_YEAR]["wmghg"] / agg_b[REF_YEAR]["wmghg"]
    wm_hi = agg_h[REF_YEAR]["wmghg"] / agg_b[REF_YEAR]["wmghg"]
    print(f"   aerosol 2019 ratios p05/best {aer_ratio_lo:.3f}, p95/best "
          f"{aer_ratio_hi:.3f}   (workbook: x1.67 / x0.41)")
    print(f"   WMGHG   2019 ratios p05/best {wm_lo:.3f}, p95/best "
          f"{wm_hi:.3f}   (workbook: x0.91 / x1.09)")

    def scaled(factors):
        return sum(series[c] * factors.get(c, 1.0) for c in COMPONENTS)

    for lab, fac, refs in (
            ("aerosol x 1.67", {c: aer_ratio_lo for c in AEROSOL}, ("AB40", "AC40")),
            ("aerosol x 0.41", {c: aer_ratio_hi for c in AEROSOL}, ("AB41", "AC41"))):
        b = slope(anom, scaled(fac), oni)
        ck(f"{lab}: b", b, cell(refs[0]), 5e-4)
        ck("   b * F2x", b * F2X, cell(refs[1]), 5e-4, "{:.3f}")
    bw_lo = slope(anom, scaled({c: wm_lo for c in WMGHG}), oni)
    bw_hi = slope(anom, scaled({c: wm_hi for c in WMGHG}), oni)
    print(f"   WMGHG x 0.91 / x 1.09: b = {bw_lo:.4f} / {bw_hi:.4f}"
          f"   workbook AB42 '{S['AB42'].value}'")
    print(f"   WMGHG x 0.91 / x 1.09: b * F2x = {bw_lo * F2X:.3f} / "
          f"{bw_hi * F2X:.3f}   workbook AC42 '{S['AC42'].value}'")

    print("\n4. Single-component ranges, b * F2x   [Sensitivity!AC43]")
    singles = {}
    for c in COMPONENTS:
        lo = agg_l[REF_YEAR][c] / agg_b[REF_YEAR][c]
        hi = agg_h[REF_YEAR][c] / agg_b[REF_YEAR][c]
        v = sorted((slope(anom, scaled({c: lo}), oni) * F2X,
                    slope(anom, scaled({c: hi}), oni) * F2X))
        singles[c] = v
    for c, v in sorted(singles.items(), key=lambda kv: kv[1][0] - kv[1][1]):
        print(f"   {c:34} {v[0]:.2f} - {v[1]:.2f}")
    print(f"   workbook AC43: '{S['AC43'].value}'")

    print(f"\n5. Monte Carlo, {draws:,} draws, seed {seed}"
          f"   [Sensitivity!AC44:AD46]")
    rel_sd = {}
    for c in COMPONENTS:
        half = (agg_h[REF_YEAR][c] - agg_l[REF_YEAR][c]) / 2.0 / Z90
        rel_sd[c] = abs(half / agg_b[REF_YEAR][c])
    print("   relative sd of each component's scale factor (2019 half-width / 1.645):")
    for c in COMPONENTS:
        print(f"      {c:34} {rel_sd[c]*100:5.1f} %")

    rng = np.random.default_rng(seed)
    mat = np.column_stack([series[c] for c in COMPONENTS])
    sd_vec = np.array([rel_sd[c] for c in COMPONENTS])
    i_co2 = COMPONENTS.index("CO2")
    A0 = np.column_stack([np.ones(len(years)), np.zeros(len(years)), oni])

    def run(scale_co2=True, tie_f2x=False, add_noise=False):
        f = 1.0 + rng.standard_normal((draws, len(COMPONENTS))) * sd_vec
        if not scale_co2:
            f[:, i_co2] = 1.0
        out = np.empty(draws)
        noise = (rng.standard_normal(draws) * float(F["U19"].value)
                 if add_noise else np.zeros(draws))
        for k in range(draws):
            A = A0.copy()
            A[:, 1] = mat @ f[k]
            beta, *_ = np.linalg.lstsq(A, anom, rcond=None)
            f2x = F2X * f[k, i_co2] if tie_f2x else F2X
            out[k] = (beta[1] + noise[k]) * f2x
        return out

    cached = {}
    for lab, kw, refs in (
            ("all 11 components, F2x fixed", dict(), ("AC44", "AD44")),
            ("CO2 held at best (non-CO2 only)", dict(scale_co2=False), ("AC45", "AD45")),
            ("full: F2x tied to the CO2 draw + AR(1) noise",
             dict(tie_f2x=True, add_noise=True), ("AC46", "AD46"))):
        d = run(**kw)
        cached[refs[0]] = d
        p5, p50, p95 = np.percentile(d, [5, 50, 95])
        print(f"   {lab}")
        print(f"      5th / 50th / 95th {p5:.2f} / {p50:.2f} / {p95:.2f}"
              f"      workbook {S[refs[0]].value}")
        print(f"      sd {d.std(ddof=1):.3f}, 2.5-97.5% "
              f"{np.percentile(d, 2.5):.2f}-{np.percentile(d, 97.5):.2f}"
              f"      workbook {S[refs[1]].value}")

    print("\n6. Quadrature comparison   [Sensitivity!AC47:AD47]")
    se_ts = float(F["U19"].value) * F2X
    se_f2x = b_best * (F2X_HALFWIDTH_5_95 / Z90)
    se_nonco2 = cached["AC45"].std(ddof=1)
    tot = (se_ts ** 2 + se_f2x ** 2 + se_nonco2 ** 2) ** 0.5
    print(f"   time series {se_ts:.3f} (workbook 0.066), F2x {se_f2x:.3f} "
          f"(0.117), non-CO2 ERF {se_nonco2:.3f} (0.068)")
    ck("total SE in quadrature", tot, cell("AC47"), 5e-3, "{:.2f}")
    c = b_best * F2X
    print(f"   90% {c - Z90*tot:.2f}-{c + Z90*tot:.2f}, 95% "
          f"{c - 1.96*tot:.2f}-{c + 1.96*tot:.2f}"
          f"      workbook '{S['AD47'].value}'")
    print("\n   Note the point made at Sensitivity!AA51: the quadrature row double-counts,"
          "\n   because CO2-ERF and F2x uncertainty are the same physical quantity. The tied"
          "\n   Monte Carlo in section 5 is the defensible interval.")


if __name__ == "__main__":
    a = sys.argv[1:]
    if len(a) < 2:
        sys.exit("usage: erf_uncertainty.py <workbook.xlsx> <erf_csv_directory> "
                 "[draws] [seed]")
    main(a[0], a[1],
         int(a[2]) if len(a) > 2 else DRAWS,
         int(a[3]) if len(a) > 3 else SEED)
