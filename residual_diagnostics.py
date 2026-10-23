#!/usr/bin/env python3
"""
residual_diagnostics.py

Heteroskedasticity and variance-clustering diagnostics for the regression
residuals in Climate_Change_Analysis_2.xlsx.

Added September/October 2026 in response to a multi-model review of the
workbook, which noted that the workbook reported autocorrelation diagnostics
(lag-1, AR(1)-inflated SEs, Newey-West cross-check) but no formal
heteroskedasticity test, even though climate residuals are expected to show
variance clustering around major eruptions and strong ENSO events.

Models tested
-------------
  annual_co2 Annual tab, rows 2-68:   mean anomaly ~ log2(CO2/280)
            residual column Annual!H, regressor Annual!E
  monthly   Volcanic tab, rows 9-822:  anomaly ~ log2(CO2/280) + ONI(t-3) + AOD(t-7)
            residual column Volcanic!J, regressors Volcanic!C:E
  kernel    Kernel tab, rows 9-815:    anomaly ~ response-filtered CO2, ONI, AOD
            residual column Kernel!M, regressors Kernel!H:J
  annual    Forcing tab, rows 2-68:    anomaly ~ anthropogenic ERF + ONI
            residual column Forcing!Q, regressors Forcing!AN:AO

Tests
-----
  Breusch-Pagan (classic, assumes Gaussian residuals)
  Breusch-Pagan (Koenker studentized form, robust to non-normality)
  White (squares and cross-products; general form)
  ARCH-LM at q = 1, 6, 12 (variance clustering in time)
  The same tests on the AR(1)-filtered residual u(t) = e(t) - rho*e(t-1), which is
    what Diagnostics rows 17-19 report and where the headline result lives: the
    large raw ARCH statistics collapse once the known level autocorrelation is
    removed, so they were an artefact of it rather than evidence of clustering
  Sub-period and conditional residual standard deviations
  Slope standard errors: naive OLS, HC3 (heteroskedasticity-consistent),
    Newey-West (heteroskedasticity- AND autocorrelation-consistent)

Usage
-----
  python3 residual_diagnostics.py [path_to_workbook.xlsx]

Dependencies: numpy, scipy, openpyxl.

This script computes every live statistic on the workbook's Diagnostics tab:
sections A (including the AR(1)-filtered rows 16-19), the naive and AR(1)-inflated
standard errors of section B, and the subset spreads and variance ratios of
section C. The HC3 and Newey-West figures in section B are the six cells the tab
shows in blue, because they need matrix operations that are impractical as
worksheet formulas; everything else there should match cell for cell.

Note on interpretation: the headline intervals in the workbook are built from
AR(1)-inflated and Newey-West standard errors. Newey-West is HAC, so it is
already robust to the heteroskedasticity these tests detect. The tests are
therefore diagnostic (do the residuals behave as OLS assumes?) rather than
corrective (do the published intervals need to change?).
"""

import sys
import numpy as np
from scipy import stats
import openpyxl

DEFAULT_WB = "Climate_Change_Analysis_2.xlsx"

MODELS = {
    "monthly": dict(sheet="Volcanic", r0=9, r1=822, y=10, X=[3, 4, 5], ycol=2,
                    names=["log2(CO2/280)", "ONI(t-3)", "AOD(t-7)"],
                    label="Monthly three-factor, single lags (Volcanic!J)"),
    "kernel":  dict(sheet="Kernel", r0=9, r1=815, y=13, X=[8, 9, 10], ycol=3,
                    names=["resp log2(CO2)", "resp ONI", "resp AOD"],
                    label="Monthly kernel model, exponential responses (Kernel!M)"),
    "annual":  dict(sheet="Forcing", r0=2, r1=68, y=17, X=[40, 41], ycol=2,
                    names=["anthropogenic ERF", "ONI"],
                    label="Annual anthropogenic-ERF model (Forcing!Q)"),
    "annual_co2": dict(sheet="Annual", r0=2, r1=68, y=8, X=[5], ycol=4,
                       names=["log2(CO2/280)"],
                       label="Annual CO2-proxy model, the published 2.51-2.98 interval (Annual!H)"),
}


def column(ws, col, r0, r1):
    out = []
    for r in range(r0, r1 + 1):
        v = ws.cell(r, col).value
        out.append(float(v) if type(v) in (int, float) else np.nan)
    return np.array(out, dtype=float)


def ols(y, X):
    """X without intercept; returns beta (incl. intercept), resid, R^2, XtX_inv."""
    A = np.column_stack([np.ones(len(y)), X])
    beta, *_ = np.linalg.lstsq(A, y, rcond=None)
    resid = y - A @ beta
    ss_tot = ((y - y.mean()) ** 2).sum()
    r2 = 1.0 - (resid ** 2).sum() / ss_tot if ss_tot > 0 else np.nan
    return beta, resid, r2, np.linalg.inv(A.T @ A), A


def lm_test(y, X, df):
    """Lagrange-multiplier test LM = n * R^2 of auxiliary regression."""
    _, _, r2, _, _ = ols(y, X)
    n = len(y)
    lm = n * r2
    return lm, df, stats.chi2.sf(lm, df), r2, n


def breusch_pagan(e, X):
    """Classic (Gaussian) and Koenker (studentized) forms."""
    n = len(e)
    e2 = e ** 2
    sigma2 = e2.mean()
    # Koenker: LM = n * R^2 from e^2 on X
    lm_k, df, p_k, r2, _ = lm_test(e2, X, X.shape[1])
    # Classic: LM = 0.5 * explained sum of squares of (e^2 / sigma2) on X
    g = e2 / sigma2
    _, resid_g, _, _, _ = ols(g, X)
    ess = ((g - g.mean()) ** 2).sum() - (resid_g ** 2).sum()
    lm_c = 0.5 * ess
    return (lm_c, df, stats.chi2.sf(lm_c, df)), (lm_k, df, p_k, r2)


def white_test(e, X):
    """General White test: e^2 on levels, squares and cross-products."""
    n, k = X.shape
    cols = [X]
    cols.append(X ** 2)
    for i in range(k):
        for j in range(i + 1, k):
            cols.append((X[:, i] * X[:, j]).reshape(-1, 1))
    Z = np.column_stack(cols)
    # A dropped column would change the degrees of freedom and silently stop
    # matching the workbook's hard-coded df, so refuse rather than adjust.
    keep = [c for c in range(Z.shape[1]) if np.ptp(Z[:, c]) > 0]
    if len(keep) != Z.shape[1]:
        raise ValueError(
            f"White design is degenerate: {Z.shape[1] - len(keep)} constant "
            f"column(s). The workbook's df would no longer apply.")
    lm, df, p, r2, _ = lm_test(e ** 2, Z, Z.shape[1])
    return lm, df, p, r2


def arch_lm(e, q):
    """Engle ARCH-LM: e^2_t on q of its own lags."""
    e2 = e ** 2
    n = len(e2)
    y = e2[q:]
    X = np.column_stack([e2[q - i: n - i] for i in range(1, q + 1)])
    lm, df, p, r2, nn = lm_test(y, X, q)
    return lm, df, p, r2, nn


def hc3_se(X, e, XtX_inv, A):
    n = len(e)
    h = np.einsum("ij,jk,ik->i", A, XtX_inv, A)
    w = (e / (1.0 - h)) ** 2
    meat = A.T @ (A * w[:, None])
    V = XtX_inv @ meat @ XtX_inv
    return np.sqrt(np.diag(V))


def newey_west_se(A, e, XtX_inv, lags):
    n = len(e)
    S = (A * e[:, None]).T @ (A * e[:, None])
    for L in range(1, lags + 1):
        w = 1.0 - L / (lags + 1.0)
        G = (A[L:] * e[L:, None]).T @ (A[:-L] * e[:-L, None])
        S += w * (G + G.T)
    V = XtX_inv @ S @ XtX_inv
    return np.sqrt(np.diag(V))


def lag1(x):
    return np.corrcoef(x[:-1], x[1:])[0, 1]


def filtered_tests(e, X, lags=(1, 6)):
    """Diagnostics rows 16-19: prewhiten with the lag-1 coefficient, then retest.

    An AR(1) process in the levels makes the squares autocorrelated too, so
    ARCH-LM over-rejects on residuals with lag-1 near 0.5. Removing that
    structure first is what separates genuine conditional heteroskedasticity
    from the autocorrelation the workbook already documents and corrects for.
    """
    rho = lag1(e)
    u = e[1:] - rho * e[:-1]
    out = {"rho": rho, "u": u, "n": len(u), "lags": tuple(lags)}
    lm, df, p, _, _ = lm_test(u ** 2, X[1:], X.shape[1])
    out["bp"] = (lm, df, p, len(u))
    for q in lags:
        out[f"arch{q}"] = arch_lm(u, q)
    return out


def report(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    print(f"Workbook: {path}\n")
    results = {}

    for key, m in MODELS.items():
        ws = wb[m["sheet"]]
        e = column(ws, m["y"], m["r0"], m["r1"])
        X = np.column_stack([column(ws, c, m["r0"], m["r1"]) for c in m["X"]])
        ok = np.isfinite(e) & np.isfinite(X).all(axis=1)
        e, X = e[ok], X[ok]
        n = len(e)

        print("=" * 78)
        print(f"{key.upper()}  {m['label']}")
        print(f"rows {m['r0']}-{m['r1']} of {m['sheet']}, n = {n}, "
              f"residual sd = {e.std(ddof=1):.6f} degC, lag-1 = {lag1(e):.4f}")
        print("-" * 78)

        (lmc, dfc, pc), (lmk, dfk, pk, r2k) = breusch_pagan(e, X)
        lmw, dfw, pw, r2w = white_test(e, X)
        print(f"Breusch-Pagan, classic      LM = {lmc:9.3f}  df {dfc}  p = {pc:.3e}")
        print(f"Breusch-Pagan, Koenker      LM = {lmk:9.3f}  df {dfk}  p = {pk:.3e}   (aux R2 {r2k:.4f})")
        print(f"White, squares + cross      LM = {lmw:9.3f}  df {dfw}  p = {pw:.3e}   (aux R2 {r2w:.4f})")
        for q in (1, 6, 12):
            if n > 4 * q:
                lma, dfa, pa, r2a, na = arch_lm(e, q)
                print(f"ARCH-LM, q = {q:<2d}              LM = {lma:9.3f}  df {dfa}  p = {pa:.3e}   (n {na})")

        if key in ("monthly", "annual"):
            ft = filtered_tests(e, X, lags=(1,) if key == "annual" else (1, 6))
            print("-" * 78)
            print("the same tests after removing the level autocorrelation   "
                  "[Diagnostics rows 16-19]")
            print(f"rho, lag-1 of e (Diagnostics!AX16)   {ft['rho']:.6f}   "
                  f"u sd {ft['u'].std(ddof=1):.6f}, lag-1 of u "
                  f"{lag1(ft['u']):+.4f}")
            lm, df, p, nn = ft["bp"]
            print(f"Breusch-Pagan, Koenker, on u LM = {lm:9.3f}  df {df}  "
                  f"p = {p:.3e}   (n {nn})")
            for q in ft["lags"]:
                lma, dfa, pa, _, na = ft[f"arch{q}"]
                print(f"ARCH-LM on u, q = {q:<2d}         LM = {lma:9.3f}  df {dfa}  "
                      f"p = {pa:.3e}   (n {na})")
            if key == "monthly":
                print("the raw ARCH statistics above were an artefact of the level "
                      "autocorrelation")
            else:
                print("the significant raw ARCH q=1 above is a single-event effect "
                      "(the 1992-93 Pinatubo residuals), not the mechanism of the "
                      "monthly block: a pure AR(1) artefact at this rho would "
                      "predict LM near 1.1")
            results[f"{key}_filtered"] = ft

        results[key] = dict(n=n, sd=e.std(ddof=1), lag1=lag1(e),
                            bp_c=(lmc, dfc, pc), bp_k=(lmk, dfk, pk),
                            white=(lmw, dfw, pw))
        print()

    # Slope standard errors for the two headline regressions
    print("=" * 78)
    print("HEADLINE SLOPE STANDARD ERRORS")
    print("-" * 78)
    for key, slope_idx, nw_lags in (("annual_co2", 1, 6), ("annual", 1, 6),
                                    ("monthly", 1, 12)):
        m = MODELS[key]
        ws = wb[m["sheet"]]
        y = column(ws, m["ycol"], m["r0"], m["r1"])
        X = np.column_stack([column(ws, c, m["r0"], m["r1"]) for c in m["X"]])
        ok = np.isfinite(y) & np.isfinite(X).all(axis=1)
        y, X = y[ok], X[ok]
        beta, e, r2, XtX_inv, A = ols(y, X)
        n, k = A.shape
        s2 = (e ** 2).sum() / (n - k)
        se_ols = np.sqrt(np.diag(XtX_inv) * s2)
        se_hc3 = hc3_se(X, e, XtX_inv, A)
        se_nw = newey_west_se(A, e, XtX_inv, nw_lags)
        rho = lag1(e)
        infl = np.sqrt((1 + rho) / (1 - rho))
        i = slope_idx
        print(f"{key}: slope on {m['names'][0]} = {beta[i]:.6f}  (R2 {r2:.4f}, n {n})")
        print(f"   naive OLS SE          {se_ols[i]:.6f}")
        print(f"   HC3 robust SE         {se_hc3[i]:.6f}   ratio to naive {se_hc3[i]/se_ols[i]:.3f}")
        print(f"   Newey-West SE ({nw_lags} lags) {se_nw[i]:.6f}   ratio to naive {se_nw[i]/se_ols[i]:.3f}")
        if key == "monthly":
            # v1.6: Newey-West has not converged at 12 lags on the monthly residuals
            long = [(L, newey_west_se(A, e, XtX_inv, L)[i]) for L in (24, 36, 60, 96)]
            print("   Newey-West at longer lags: "
                  + ", ".join(f"{L}:{v:.4f}" for L, v in long))
            print(f"   converged value about {long[-1][1]:.3f}, "
                  f"{long[-1][1]/(se_ols[i]*infl):.1f}x the AR(1)-inflated figure, "
                  f"so the AR(1) factor understates the dependence here")
        span = [newey_west_se(A, e, XtX_inv, L)[i] for L in range(1, nw_lags + 1)]
        print(f"   Newey-West, 1-{nw_lags} lags{'' if nw_lags > 9 else ' '}  "
              f"{min(span):.6f} to {max(span):.6f}")
        print(f"   AR(1)-inflated SE     {se_ols[i]*infl:.6f}   (rho {rho:.4f}, factor {infl:.3f})")
        print()

    # Where the variance clusters: monthly three-factor residuals
    print("=" * 78)
    print("WHERE THE VARIANCE CLUSTERS  (monthly three-factor residuals)")
    print("-" * 78)
    ws = wb["Volcanic"]
    e = column(ws, 10, 9, 822)
    yr = column(ws, 12, 9, 822)
    oni = column(ws, 4, 9, 822)
    aod = column(ws, 7, 9, 822)          # same-month stratospheric AOD
    ok = np.isfinite(e) & np.isfinite(yr)
    e, yr, oni, aod = e[ok], yr[ok], oni[ok], aod[ok]
    full = e.std(ddof=1)
    print(f"all months                      n {len(e):4d}   sd {full:.4f}")
    epochs = [(1958, 1970), (1971, 1985), (1986, 2000), (2001, 2015), (2016, 2027)]
    for a, b in epochs:
        s = (yr >= a) & (yr < b + 1)
        if s.sum() > 2:
            print(f"{a}-{min(b,2026)}                       n {s.sum():4d}   sd {e[s].std(ddof=1):.4f}"
                  f"   ratio {e[s].std(ddof=1)/full:.2f}")
    print()
    hi = np.isfinite(aod) & (aod > 0.01)
    lo = np.isfinite(aod) & (aod <= 0.01)
    print(f"stratospheric AOD > 0.01        n {hi.sum():4d}   sd {e[hi].std(ddof=1):.4f}"
          f"   ratio {e[hi].std(ddof=1)/full:.2f}")
    print(f"stratospheric AOD <= 0.01       n {lo.sum():4d}   sd {e[lo].std(ddof=1):.4f}"
          f"   ratio {e[lo].std(ddof=1)/full:.2f}")
    # v1.6: the 2016-2026 subset straddles the 2023 step, so its spread about its own
    # mean measures the shift; about the two sub-period means it is slightly narrower
    recent = yr >= 2016
    pre, post = recent & (yr < 2023), recent & (yr >= 2023)
    pooled = np.concatenate([e[pre] - e[pre].mean(), e[post] - e[post].mean()])
    sd_sub = np.sqrt((pooled ** 2).sum() / (pre.sum() + post.sum() - 2))
    print(f"2016-2026 about its own mean    n {recent.sum():4d}   sd "
          f"{e[recent].std(ddof=1):.4f}   ratio {e[recent].std(ddof=1)/full:.2f}")
    print(f"   about the two sub-period means          sd {sd_sub:.4f}   "
          f"ratio {sd_sub/full:.2f}   (means {e[pre].mean():+.4f} and "
          f"{e[post].mean():+.4f})")
    print()
    strong = np.isfinite(oni) & (np.abs(oni) >= 1.0)
    neutral = np.isfinite(oni) & (np.abs(oni) < 0.5)
    print(f"|ONI| >= 1.0 (strong ENSO)      n {strong.sum():4d}   sd {e[strong].std(ddof=1):.4f}"
          f"   ratio {e[strong].std(ddof=1)/full:.2f}")
    print(f"|ONI| < 0.5 (neutral)           n {neutral.sum():4d}   sd {e[neutral].std(ddof=1):.4f}"
          f"   ratio {e[neutral].std(ddof=1)/full:.2f}")
    f_stat = e[hi].var(ddof=1) / e[lo].var(ddof=1)
    p_f = 2 * min(stats.f.sf(f_stat, hi.sum() - 1, lo.sum() - 1),
                  stats.f.cdf(f_stat, hi.sum() - 1, lo.sum() - 1))
    print(f"\nvariance ratio, high vs low AOD  F = {f_stat:.3f}  p = {p_f:.3e} (two-sided)")
    h1 = yr < 1992
    f2 = e[h1].var(ddof=1) / e[~h1].var(ddof=1)
    p2 = 2 * min(stats.f.sf(f2, h1.sum() - 1, (~h1).sum() - 1),
                 stats.f.cdf(f2, h1.sum() - 1, (~h1).sum() - 1))
    print(f"variance ratio, pre/post 1992    F = {f2:.3f}  p = {p2:.3e} (two-sided)")
    print("\n" + "=" * 78)
    print("SUMMARY: what each model's residuals show")
    print("-" * 78)
    for key in ("monthly", "kernel", "annual", "annual_co2"):
        if key not in results:
            continue
        r = results[key]
        print(f"   {key:<11} n {r['n']:>4}  sd {r['sd']:.4f}  lag-1 {r['lag1']:+.3f}"
              f"   BP(Koenker) p {r['bp_k'][2]:.3f}   White p {r['white'][2]:.3f}")
    if "annual_filtered" in results:
        fa = results["annual_filtered"]
        print(f"   annual ERF, after removing the level autocorrelation: "
              f"ARCH q=1 p {fa['arch1'][2]:.3f} (raw p 0.007)")
    if "monthly_filtered" in results:
        ft = results["monthly_filtered"]
        print(f"   monthly, after removing the level autocorrelation: "
              f"BP p {ft['bp'][2]:.3f}, "
              + ", ".join(f"ARCH q={q} p {ft[f'arch{q}'][2]:.3f}" for q in ft["lags"]))
    return results


if __name__ == "__main__":
    report(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_WB)
