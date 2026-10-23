# Licence for the workbook and documentation

Copyright (c) 2026 Russ Nida.

The workbook `Climate_Change_Analysis_2.xlsx`, its two Markdown exports, and
`HANDOFF.md` are licensed under a
[Creative Commons Attribution 4.0 International Licence](https://creativecommons.org/licenses/by/4.0/)
(CC BY 4.0).

You are free to share and adapt this material, including for commercial
purposes, provided you give appropriate credit, link to the licence, and
indicate if changes were made. There is no NonCommercial restriction and no
ShareAlike requirement.

The full legal text is at <https://creativecommons.org/licenses/by/4.0/legalcode>.
This file states the licence rather than reproducing it, which is the marking
Creative Commons itself recommends; if you would rather the repository were
self-contained, the legal code can be added verbatim.

Suggested citation:

> Nida, R. (2026). *Climate_Change_Analysis_2*: global temperature against CO2
> with natural-variability controls. Version 1.7.11. CC BY 4.0.

## What this licence does not cover

### The code

Every Python script in this repository is under the MIT Licence, in `LICENSE`,
including any added later. Creative Commons advises against using CC licences for
software, so the code carries a software licence of its own.

### The input data

Nothing here places any licence on the input datasets, and for most of them no
one could. The authority is the **Terms** column of the workbook's `Sources`
tab, item by item. In summary:

| Input | Terms |
|---|---|
| NASA GISTEMP v4 surface temperature | US Government work, public domain; citation requested |
| NOAA GML Mauna Loa CO2 (item 2) | Free; NOAA GML requests citation. Cite both NOAA GML and Keeling et al. (1976): March 1958 to April 1974 is Scripps data, and this series starts in March 1958 |
| NOAA Solar CDR v2.1 (NRLTSI2) and v3 (NNLTSI1), items 3 and 4 | NOAA CDR use agreement. Resolved v1.7.11 against UseAgreement_01B-32.pdf: non-proprietary, publicly available, no restrictions on use; citation urged, not required |
| NOAA CPC ONI on ERSSTv5 and ERSSTv6 | US Government work, public domain |
| NOAA CPC RONI v5 (item 15) | US Government work, public domain. Resolved v1.7.10: CPC requests no citation, per NWS PNS 26-05. Note there are two RONI series; this is the ERSSTv5 one |
| ClimateIndicator forcing-timeseries (items 7, 9, 10) | MIT, Copyright (c) 2023 Chris Smith. Verified against the clone at commit 852f4f7. Notices must travel with copies: see `THIRD-PARTY-NOTICES` |
| GloSSAC stratospheric aerosol | NASA data, public domain; cite GloSSAC |
| CERES EBAF Ed4.2.1 (item 13) | NASA data, openly available, citation requested. Recorded v1.7.11; cite Loeb et al. (2018) and the product DOI |
| NOAA NCEI ocean heat content (item 14) | US Government work, public domain, citation requested. Recorded v1.7.11; cite Levitus et al. (2012) |
| HadCRUT5 | Open Government Licence v3 |
| Berkeley Earth | CC BY 4.0 |

Four of those carry requirements that survive any onward reuse of this work. The
ClimateIndicator repository is MIT, which is the strictest in form: the copyright
notice and the permission notice must be included with copies, which a citation
line does not satisfy, so they are reproduced in full in `THIRD-PARTY-NOTICES`.
HadCRUT5 (Open Government Licence v3) and Berkeley Earth (CC BY 4.0) are
satisfied by attribution. Reusing this repository under CC BY 4.0 discharges none
of them.

Beyond those four, NASA and NOAA request citation of GISTEMP, GloSSAC and the
Mauna Loa record, and the two solar irradiance series are governed by the NOAA
CDR use agreement rather than by a licence that permits unrestricted
redistribution.

## Before publishing this repository

**Nothing is outstanding as of v1.7.11.** Every item this file listed has been
resolved against a primary source, and each closure is recorded in the workbook
rather than only here:

| Was open | Closed |
|---|---|
| ClimateIndicator repository licence | v1.7.9, against the clone at commit 852f4f7: MIT, not CC BY 4.0 |
| RONI provenance and citation | v1.7.10, against a checksummed download and the publisher record |
| Literature entries drafted from memory | v1.7.11, all against Crossref and DataCite; one correction, at `L8` |
| Items 13 and 14 with no terms recorded | v1.7.11, location, terms and citation written for both |
| NOAA CDR terms for items 3 and 4 | v1.7.11, against the agency's own use agreement: no restrictions on use |
| Whether the IGCC 2025 ERF values were preliminary | v1.7.11, against the published paper: they are the final values |

Two things are recorded rather than reconciled, and neither is a verification gap.
The aerosol percentile range in `erf_inputs` differs from the paper's by 0.02 and
0.01, which is a pointwise envelope against an ensemble quantile; see
`Sources!L13`. And `Sources` row 36 is a placeholder citing nothing.

What is left is a decision rather than a check: whether to publish, and whether
public or private.

The `Change log` tab records every revision, and `Sources` section F states the
terms inside the workbook itself so that a copy separated from this repository
still carries them.
