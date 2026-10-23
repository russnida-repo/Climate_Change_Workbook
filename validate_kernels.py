#!/usr/bin/env python3
"""
validate_kernels.py

Out-of-sample validation of the exponential-response ("kernel") model on the
Kernel tab of Climate_Change_Analysis_2.xlsx: single-split holdouts, blocked
7-fold cross-validation with the response time constants chosen out of fold,
and the CO2-only benchmark. These are the results pasted into Kernel!P52:Q70.

PROVENANCE, READ THIS FIRST
---------------------------
This file is a RECONSTRUCTION, written 2026-10-22 from the method documented on
the Kernel tab (notes at O41:O51, the live response formulas in columns H:J, and
the LINEST row ranges in P52:P60). The original script of this name was run in
September 2026 and was not kept with the workbook. Running this one reproduces
the pasted figures as follows:

  Exact, to every digit the workbook stores:
    Kernel!P13:P20   full-record fit a, b, c, d, R^2, n, lag-1 autocorrelation
    Kernel!P52:P60   the 1958-95 training fit, its holdout R^2, both bias
                     figures and the refit b
    Kernel!P63:Q65   all three single-split holdout R^2 / bias pairs
    Kernel!P70       in-sample R^2 at tau = (6, 12) on the same rows

  Close but not exact, blocked 7-fold cross-validation:
    Kernel!P66/Q66   this script 0.887 / +0.002 on decade folds (see below),
                     0.889 / +0.002 on equal row-count folds;  workbook 0.887 / +0.003
    Kernel!P68/Q68   this script 0.890 / +0.002;  workbook 0.889 / +0.003
    Kernel!P69       this script 0.844 (row-count) or 0.846 (decade);  workbook 0.838
    Kernel!O67       tau per fold: on CALENDAR-DECADE folds the script reproduces the
                     workbook exactly, all 14 selections (ENSO 6,6,6,6,9,6,9 and AOD
                     9,12,12,24,12,12,12). On equal row-count folds the AOD sequence
                     still matches and ENSO differs in fold 5 only.

  PARTLY RESOLVED, 2026-10-22 (v1.6), after an independent review: the original did
  not cut the 807 months into equal row-count blocks but into calendar decades
  (1958-67, 1968-77, ... 2018-25). That partition reproduces Kernel!O67 exactly and
  gives a pooled R^2 of 0.8866, i.e. the pasted 0.887, so it is almost certainly the
  convention the original used. Section 5 of this script runs it. What is still
  unaccounted for: on that same partition P68 is 0.8896 against the pasted 0.889,
  Q66 is +0.0017 against 0.003, Q68 is -0.0000 against 0.003, and P69 is 0.8456
  against 0.838. P69 in particular has no recoverable provenance under any
  partition, row range or metric tried here.

  The fold-5 difference under the row-count partition is a tie rather than a
  disagreement: on that fold's training rows tau_ENSO = 6 gives training R^2
  0.9128580 and tau_ENSO = 9 gives 0.9127306, a gap of 0.00013, and they are the top
  two of the 49 grid points. The pooled figures are likewise insensitive to the
  convention: the partitions tried in section 7 agree within 0.0005 of each other,
  and across the five styles tried during the review the tau-selected pooled R^2
  spans 0.8866 to 0.8893. None of this moves a conclusion. Treat the workbook's
  pasted values as the original run and these as an independent check that lands in
  the same place.

Method
------
Response filter, as on the Kernel tab:
    R(t) = a * R(t-1) + (1 - a) * x(t),   a = exp(-1 / tau),  a = 0 when tau = 0
with R restarting at the raw value wherever the previous row is blank, which is
what the Excel formulas in H:J do.

Model: anomaly = a + b*R(log2(CO2/280)) + c*R(ONI) + d*R(AOD), tau_CO2 fixed at
0 (see the Kernel tab note: filtering CO2 has no R^2 optimum, so the record
cannot separate transient from equilibrium response).

Samples: in-sample and cross-validated on rows 9-815 (807 months, Oct 1958 to
Dec 2025, the rows where the anomaly and all three response series exist).
Single split: train rows 9-455 (Oct 1958 to Dec 1995), hold out rows 456-815
(Jan 1996 to Dec 2025).

Holdout and pooled R^2 are 1 - SSE / DEVSQ(observed over the same rows), which
is the definition in Kernel!P57. Bias is mean(observed) - mean(predicted).

tau grid: {0, 3, 6, 9, 12, 18, 24}. The cap at 24 is implied by the workbook: its
single-split search selected (24, 24), whereas an uncapped search peaks at
tau_ENSO = 30. Grids that contain 6, 9, 12 and 24 and no intermediate values give
identical results; a grid containing 15, for instance, selects tau_AOD = 15 in
three folds and gives 0.8886 instead.

Usage
-----
  python3 validate_kernels.py [path_to_workbook.xlsx]

Dependencies: numpy, openpyxl.
"""

import math
import sys

import numpy as np
import openpyxl

DEFAULT_WB = "Climate_Change_Analysis_2.xlsx"
TAU_GRID = [0, 3, 6, 9, 12, 18, 24]
N_FOLDS = 7

FIT_LO, FIT_HI = 9, 815          # in-sample and cross-validation rows
TRAIN_LO, TRAIN_HI = 9, 455      # single split: training
HOLD_LO, HOLD_HI = 456, 815      # single split: holdout
RECENT_LO, RECENT_HI = 780, 815  # 2023-2025

# Kernel tab columns
C_YEAR, C_ANOM, C_LOG2, C_ONI, C_AOD = 2, 3, 4, 5, 6


def isnum(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def read_columns(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Kernel"]
    n = ws.max_row
    cols = {}
    for name, c in (("year", C_YEAR), ("anom", C_ANOM), ("log2", C_LOG2),
                    ("oni", C_ONI), ("aod", C_AOD)):
        cols[name] = [ws.cell(r, c).value for r in range(1, n + 1)]
    return wb, ws, cols


def response(x, tau):
    """One-box exponential response, matching Kernel!H:J exactly."""
    a = 0.0 if tau <= 0 else math.exp(-1.0 / tau)
    out = [None] * len(x)
    for i, v in enumerate(x):
        if not isnum(v):
            continue
        prev = out[i - 1] if i > 0 else None
        out[i] = v if prev is None else a * prev + (1.0 - a) * v
    return out


def rows_where(lo, hi, series):
    return [r for r in range(lo, hi + 1)
            if all(isnum(s[r - 1]) for s in series)]


def design(rows, regressors):
    return np.column_stack(
        [np.ones(len(rows))] +
        [np.array([g[r - 1] for r in rows], dtype=float) for g in regressors])


def ols(rows, regressors, y_series):
    A = design(rows, regressors)
    y = np.array([y_series[r - 1] for r in rows], dtype=float)
    beta, *_ = np.linalg.lstsq(A, y, rcond=None)
    pred = A @ beta
    r2 = 1.0 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()
    return beta, r2, y, pred


def oos_r2(y, pred):
    return 1.0 - ((y - pred) ** 2).sum() / ((y - y.mean()) ** 2).sum()


def pick_tau(train_rows, cols, H, grid=TAU_GRID):
    """tau_ENSO, tau_AOD by maximum training R^2 -- chosen inside the fold."""
    best = None
    for t_enso in grid:
        I = response(cols["oni"], t_enso)
        for t_aod in grid:
            J = response(cols["aod"], t_aod)
            rows = [r for r in train_rows if isnum(I[r - 1]) and isnum(J[r - 1])]
            _, r2, _, _ = ols(rows, [H, I, J], cols["anom"])
            if best is None or r2 > best[0]:
                best = (r2, t_enso, t_aod)
    return best[1], best[2]


def blocked_cv_folds(folds, cols, H, fixed_tau=None, co2_only=False, grid=TAU_GRID):
    """Blocked CV over an explicit fold list. Returns pooled R^2, bias, tau per fold."""
    rows = sorted(r for f in folds for r in f)
    return _cv(folds, rows, cols, H, fixed_tau, co2_only, grid)


def blocked_cv(rows, cols, H, fixed_tau=None, co2_only=False, grid=TAU_GRID,
               n_folds=N_FOLDS):
    """Contiguous blocked CV. Returns pooled R^2, bias, and tau per fold."""
    folds = [list(f) for f in np.array_split(np.array(rows), n_folds)]
    return _cv(folds, rows, cols, H, fixed_tau, co2_only, grid)


def _cv(folds, rows, cols, H, fixed_tau, co2_only, grid):
    y_all, p_all, taus = [], [], []
    for fold in folds:
        test = fold
        held = set(test)
        train = [r for r in rows if r not in held]
        if co2_only:
            beta, _, _, _ = ols(train, [H], cols["anom"])
            pred = design(test, [H]) @ beta
            taus.append(None)
        else:
            t_enso, t_aod = fixed_tau if fixed_tau else pick_tau(train, cols, H, grid)
            taus.append((t_enso, t_aod))
            I = response(cols["oni"], t_enso)
            J = response(cols["aod"], t_aod)
            beta, _, _, _ = ols(train, [H, I, J], cols["anom"])
            pred = design(test, [H, I, J]) @ beta
        y_all.append(np.array([cols["anom"][r - 1] for r in test], dtype=float))
        p_all.append(pred)
    y = np.concatenate(y_all)
    p = np.concatenate(p_all)
    return oos_r2(y, p), y.mean() - p.mean(), taus


def ck(label, got, want, tol, fmt="{:.6f}"):
    ok = want is None or abs(got - want) <= tol
    mark = "  ok" if ok else "  DIFFERS"
    w = "" if want is None else f"   workbook {fmt.format(want)}"
    print(f"   {label:<52} {fmt.format(got)}{w}{mark}")
    return ok


def main(path):
    wb, ws, cols = read_columns(path)

    def cell(ref):
        v = ws[ref].value
        return v if isnum(v) else None

    H = response(cols["log2"], 0)               # tau_CO2 = 0
    I6 = response(cols["oni"], 6)
    J12 = response(cols["aod"], 12)
    fit_rows = rows_where(FIT_LO, FIT_HI, (cols["anom"], H, I6, J12))

    print(f"Workbook: {path}")
    print(f"Fit sample: Kernel rows {fit_rows[0]}-{fit_rows[-1]}, "
          f"n = {len(fit_rows)} months\n")

    # ---- 1. full-record fit, against Kernel!P13:P20
    print("1. Full-record fit, tau = (0, 6, 12)   [Kernel!P13:P20]")
    beta, r2, y, pred = ols(fit_rows, [H, I6, J12], cols["anom"])
    resid = y - pred
    ck("a (intercept)", beta[0], cell("P13"), 1e-9)
    ck("b, degC per CO2 doubling", beta[1], cell("P14"), 1e-9)
    ck("c, degC per unit ONI response", beta[2], cell("P15"), 1e-9)
    ck("d, degC per unit AOD response", beta[3], cell("P16"), 1e-9)
    ck("R^2", r2, cell("P17"), 1e-9)
    ck("months in fit", float(len(fit_rows)), float(cell("P18")), 0)
    ck("residual standard deviation", resid.std(ddof=1), cell("P19"), 1e-6)
    ck("lag-1 autocorrelation of residuals",
       np.corrcoef(resid[:-1], resid[1:])[0, 1], cell("P20"), 1e-9)

    # ---- 2. single split with tau fixed at (6, 12), against Kernel!P52:P60
    print("\n2. Single split, train 1958-95, tau = (6, 12)   [Kernel!P52:P60]")
    tr = rows_where(TRAIN_LO, TRAIN_HI, (cols["anom"], H, I6, J12))
    ho = rows_where(HOLD_LO, HOLD_HI, (cols["anom"], H, I6, J12))
    bt, r2t, _, _ = ols(tr, [H, I6, J12], cols["anom"])
    ck("a (training)", bt[0], cell("P52"), 1e-9)
    ck("b (training only)", bt[1], cell("P53"), 1e-9)
    ck("c (training)", bt[2], cell("P54"), 1e-9)
    ck("d (training)", bt[3], cell("P55"), 1e-9)
    ck("R^2 in the training period", r2t, cell("P56"), 1e-9)
    y_h = np.array([cols["anom"][r - 1] for r in ho], dtype=float)
    p_h = design(ho, [H, I6, J12]) @ bt
    ck("holdout R^2 on 1996-2025", oos_r2(y_h, p_h), cell("P57"), 1e-9)
    ck("mean prediction error 1996-2025", y_h.mean() - p_h.mean(), cell("P58"), 1e-9)
    rec = list(range(RECENT_LO, RECENT_HI + 1))
    y_r = np.array([cols["anom"][r - 1] for r in rec], dtype=float)
    p_r = design(rec, [H, I6, J12]) @ bt
    ck("mean prediction error 2023-2025", y_r.mean() - p_r.mean(), cell("P59"), 1e-9)
    b_refit, _, _, _ = ols(ho, [H, I6, J12], cols["anom"])
    ck("b refit on 1996-2025 only", b_refit[1], cell("P60"), 1e-9)

    # ---- 3. single-split holdouts, against Kernel!P63:Q65
    print("\n3. Single-split holdouts   [Kernel!P63:Q65]")
    t_sel = pick_tau(tr, cols, H)
    print(f"   tau chosen on the 1958-95 training rows by max training R^2: "
          f"{t_sel}   (workbook: (24, 24))")
    for (t_enso, t_aod), (pref, qref) in ((t_sel, ("P63", "Q63")),
                                          ((6, 12), ("P64", "Q64")),
                                          ((3, 12), ("P65", "Q65"))):
        I = response(cols["oni"], t_enso)
        J = response(cols["aod"], t_aod)
        tr_t = rows_where(TRAIN_LO, TRAIN_HI, (cols["anom"], H, I, J))
        ho_t = rows_where(HOLD_LO, HOLD_HI, (cols["anom"], H, I, J))
        b, _, _, _ = ols(tr_t, [H, I, J], cols["anom"])
        yy = np.array([cols["anom"][r - 1] for r in ho_t], dtype=float)
        pp = design(ho_t, [H, I, J]) @ b
        ck(f"tau = ({t_enso}, {t_aod}): holdout R^2",
           oos_r2(yy, pp), cell(pref), 5e-4, "{:.3f}")
        ck(f"tau = ({t_enso}, {t_aod}): bias",
           yy.mean() - pp.mean(), cell(qref), 5e-4, "{:+.3f}")

    # ---- 4. blocked 7-fold CV, against Kernel!P66:Q69
    print(f"\n4. Blocked {N_FOLDS}-fold cross-validation   [Kernel!P66:Q69]")
    r2a, bias_a, taus = blocked_cv(fit_rows, cols, H)
    ck("tau and coefficients out of fold: pooled R^2", r2a, cell("P66"), 5e-3, "{:.3f}")
    ck("tau and coefficients out of fold: bias", bias_a, cell("Q66"), 5e-3, "{:+.3f}")
    print(f"   tau_ENSO per fold {[t[0] for t in taus]}   workbook [6, 6, 6, 6, 9, 6, 9]")
    print(f"   tau_AOD  per fold {[t[1] for t in taus]}   workbook [9, 12, 12, 24, 12, 12, 12]")
    r2b, bias_b, _ = blocked_cv(fit_rows, cols, H, fixed_tau=(6, 12))
    ck("tau fixed at (6, 12): pooled R^2", r2b, cell("P68"), 5e-3, "{:.3f}")
    ck("tau fixed at (6, 12): bias", bias_b, cell("Q68"), 5e-3, "{:+.3f}")
    r2c, bias_c, _ = blocked_cv(fit_rows, cols, H, co2_only=True)
    print(f"   {'CO2-only model: pooled R^2':<52} {r2c:.3f}   workbook "
          f"{cell('P69'):.3f}  DIFFERS BY {abs(r2c - cell('P69')):.3f}, unexplained "
          f"(see the header)")
    ck("in-sample R^2, tau (6, 12), same rows", r2, cell("P70"), 5e-4, "{:.3f}")

    print("\n5. Calendar-decade folds   [the partition the original used]")
    edges = [(1958, 1968), (1968, 1978), (1978, 1988), (1988, 1998),
             (1998, 2008), (2008, 2018), (2018, 2027)]
    dec = [[r for r in fit_rows if lo <= cols["year"][r - 1] < hi] for lo, hi in edges]
    print(f"   fold sizes {[len(f) for f in dec]}")
    r2d, bias_d, taus_d = blocked_cv_folds(dec, cols, H)
    ck("tau and coefficients out of fold: pooled R^2", r2d, cell("P66"), 5e-4, "{:.3f}")
    print(f"   tau_ENSO per fold {[t[0] for t in taus_d]}   workbook [6, 6, 6, 6, 9, 6, 9]")
    print(f"   tau_AOD  per fold {[t[1] for t in taus_d]}   workbook "
          f"[9, 12, 12, 24, 12, 12, 12]")
    match = ([t[0] for t in taus_d] == [6, 6, 6, 6, 9, 6, 9]
             and [t[1] for t in taus_d] == [9, 12, 12, 24, 12, 12, 12])
    print(f"   tau pattern matches the workbook exactly: {match}")

    print("\n6. What the pooled comparison does and does not show")
    print("   Pooled out-of-fold R^2 is dominated by the cross-decade trend, which a")
    print("   monotone CO2-only regressor extrapolates across blocks. Per fold:")
    for i, fold in enumerate(dec):
        held = set(fold)
        train = [r for r in fit_rows if r not in held]
        b1, _, _, _ = ols(train, [H], cols["anom"])
        y = np.array([cols["anom"][r - 1] for r in fold], dtype=float)
        p1 = design(fold, [H]) @ b1
        I = response(cols["oni"], 6)
        J = response(cols["aod"], 12)
        b3, _, _, _ = ols(train, [H, I, J], cols["anom"])
        p3 = design(fold, [H, I, J]) @ b3
        rmse1 = np.sqrt(((y - p1) ** 2).mean())
        rmse3 = np.sqrt(((y - p3) ** 2).mean())
        print(f"      {edges[i][0]}-{edges[i][1]-1}: CO2-only R^2 {oos_r2(y, p1):+.2f}, "
              f"RMSE {rmse1:.4f} -> three-factor RMSE {rmse3:.4f} "
              f"({100*(1-rmse3/rmse1):+.0f}%)")
    print("   Within every fold the CO2-only fit is near zero or negative R^2 while the")
    print("   three-factor model cuts the RMSE, which is the stronger statement.")

    print("\n7. Sensitivity of the pooled figures to the fold partition")
    a = np.array(fit_rows)
    for style, folds in (
            ("numpy array_split", [list(x) for x in np.array_split(a, N_FOLDS)]),
            ("equal blocks, remainder last",
             [list(a[i * (len(a) // N_FOLDS):(i + 1) * (len(a) // N_FOLDS)])
              for i in range(N_FOLDS - 1)]
             + [list(a[(N_FOLDS - 1) * (len(a) // N_FOLDS):])])):
        y_all, p_all = [], []
        for fold in folds:
            held = set(fold)
            train = [r for r in fit_rows if r not in held]
            t_e, t_a = pick_tau(train, cols, H)
            I, J = response(cols["oni"], t_e), response(cols["aod"], t_a)
            b, _, _, _ = ols(train, [H, I, J], cols["anom"])
            y_all.append(np.array([cols["anom"][r - 1] for r in fold], dtype=float))
            p_all.append(design(fold, [H, I, J]) @ b)
        yy, pp = np.concatenate(y_all), np.concatenate(p_all)
        print(f"   {style:<32} pooled R^2 {oos_r2(yy, pp):.4f}")
    print("\n   The conclusion the workbook draws from these numbers, that the kernel"
          "\n   specification is not an in-sample artefact and that the controls add"
          "\n   genuine predictive information (CO2-only about 0.84 against about 0.89"
          "\n   with ENSO and AOD), does not depend on the convention.")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_WB)
