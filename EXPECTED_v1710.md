# What the v1.7.10 release should contain

**Outcome, 2026-10-23: every prediction below passed.** The release hashes
`2f8d2e03`. All eight predicted values match, the structural uncertainty came out
at +0.0582, all five headline cells are bit-identical to the v1.7.4 base, and the
3,923 numeric cells that moved are all on the `Open Items` tab.

This file records, before the Excel save, the values the v1.7.10 build should
produce once Excel recalculates it. Committing to them in advance is the point:
the release is then checked against a prediction rather than against whatever it
happens to say.

Source of the predictions: `open_items.py` run against the v1.7.9 workbook with
the two-decimal `roni_v5_cpc.csv`. The script computes section A from the same
inputs the workbook uses, independently of the workbook's own formulas.

## Open Items, section A, after recalculation

| Cell | Quantity | v1.7.9 (1 decimal) | v1.7.10 expected |
|---|---|---|---|
| `AB7`  | RONI lag-3 linear trend, degC/decade | -0.0825 | **-0.0828** |
| `AB10` | ONI lag-3 linear trend, degC/decade  | -0.0048 | -0.0048 (unchanged) |
| `AB11` | correlation RONI against ONI          | 0.9519 | **0.9527** |
| `AB12` | standard deviation of RONI            | 0.8560 | **0.857** |
| `AB13` | peak of the 2023-24 event, RONI       | 1.50 | **1.49** |
| `AB18` | monthly model b, degC per doubling    | 2.8197 | **2.8203** |
| `AB24` | annual model b, degC per W/m2         | 0.4237 | **0.4236** |
| `AE24` | annual model b x F2x                  | 1.6650 | **1.6649** |

The annual block is rows 23 to 25; rows 17 to 19 are the monthly block. An
earlier draft of this table cited row 19 for the annual detrended variant, which
is the monthly one.
| `AE25` | detrended variant, b x F2x            | 1.6113 | 1.6113 (unchanged to 4dp) |

Structural uncertainty, RONI minus the ONI baseline of 1.6067:
**+0.0583 becomes +0.0582.** Still +0.06 at the precision it is reported.

## What must not move

The headline model does not use RONI, so these must be bit-identical to v1.7.4:

| Cell | Value |
|---|---|
| `Forcing!T14` | 0.4088266036401316 |
| `Forcing!T17` | 1.6066885523057173 |
| `Annual!K4`   | 2.7423885028518846 |
| `Annual!K5` / `K6` | -0.5925349218415271 / 0.9309185423756061 |

Only the `Open Items` tab should change at all. In the build, 3,082 numeric cells
differ from the v1.7.4 base and every one of them is on that tab.

## Why this file exists

The build carries the corrected inputs but stale downstream caches, which is why
it sets `fullCalcOnLoad`. A LibreOffice conversion will not fix them: tested on
this build, it recalculated `D9` and `AB18` but left `AB7`, `AB10`, `AB11`,
`AB12`, `AB13`, `AB24` and `I9` on their cached values. A LibreOffice convert is
not a recalculation of a workbook whose inputs changed outside the application,
and nothing in this package should treat it as one.

Excel with `fullCalcOnLoad` set does recalculate. So: open the build in Excel,
save, and the saved release is the canonical v1.7.10. Checking it against the
table above is the acceptance test.
