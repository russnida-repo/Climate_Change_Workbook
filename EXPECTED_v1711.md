# What the v1.7.11 release should contain

Recorded before the Excel save, so the release is checked against a prediction.
Build sha256 `5f9fe6ff`.

v1.7.11 changes text only. Every edit is in shared strings or in six cells that
were empty, and no formula, literal input or regressor was touched. The build is
rebuilt from the v1.7.4 Excel-saved base through the same single code path as
every version since v1.7.5, which is why it again carries stale caches on the
`Open Items` dependants and sets `fullCalcOnLoad`.

## The acceptance test

**The v1.7.11 release must be numerically identical to the v1.7.10 release in all
120,565 numeric cells, with zero differences.** Not "close": identical. Anything
that moved means a text edit reached a formula, which it must not have.

Checks that follow from that:

| Check | Expected |
|---|---|
| numeric cells differing from the v1.7.10 release | 0 of 120,565 |
| `Forcing!T14` | 0.4088266036401316 |
| `Forcing!T17` | 1.6066885523057173 |
| `Annual!K4` | 2.7423885028518846 |
| `Open Items!AB18` | 2.820255 |
| `Open Items!AE24` | 1.664934 |
| `Open Items!AB11` | 0.9527 |
| structural uncertainty | +0.0582 |
| `Open Items` column C against `RONI.v5.ascii.txt` | 819 rows, 0 mismatches |
| error-valued cells | none |
| worksheets | 22 |
| `Change log` entries | 20, last `v1.7.11` |

## What changed, all text

- `Sources!L8` Mauna Loa citation corrected: the previous DOI was the globally
  averaged product, and the Scripps credit the file header requests was missing.
- `Sources!L13` IGCC 2025 DOI filled in, plus a new CHECK on the preliminary
  status of the 2025 ERF values.
- `Sources!K9` and `K10` NOAA CDR terms resolved. These two cells share one
  shared string, so one edit serves both.
- `Sources!C29` IPCC Chapter 7 DOI added. The title needed no serial comma.
- `Sources!J43`, `K43`, `L43`, `J44`, `K44`, `L44` written; all six were empty.
- `Sources!L13` also records that the IGCC 2025 values were confirmed against the
  published paper, so the preliminary-data question opened and closed in the same
  revision.
- `Sources!E29` to `E34` status changed from `CHECK` to verified. These six cells
  share one shared string, so one edit serves all six; the edit replaces the whole
  `<si>` element because the bare word "CHECK" occurs nine times in sharedStrings
  and the first is inside `B3`.
- `Sources!C35` and `D35` now carry the completed Jenkins reference and its DOI.
  `D35` shared the empty-string entry with `D36`, so it was repointed rather than
  edited; `D36` must still be blank.
- `Sources!B3` standing instruction replaced with the line it asked for.
- `Sources!B48` Flanner reference marked verified.
- `Sources!B63` section F status rewritten; nothing outstanding.
- `Open Items!AA66` rewritten. It claimed every month-to-month RONI minus ONI
  difference was within 0.76; at two decimals one pair reaches 0.78, January 1992,
  RONI 2.32 against ONI 1.54. The old bound was an artifact of the rounding, since
  the one-decimal RONI read 2.3 and 2.3 minus 1.54 is exactly 0.76. The paragraph
  also still described the HTML transcription and the no-ascii-file premise.
- `Change log` gains the v1.7.11 entry.

## Package documents, updated outside the workbook

`HANDOFF.md`, `LICENSE-CONTENT.md` and `THIRD-PARTY-NOTICES` were all stale: the
first two still said version 1.7.9, the licence table still showed the RONI
citation question as open and items 13 and 14 as having no terms recorded, the
pre-publication list still had four open entries, and the notices still carried
the IGCC DOI as a CHECK. All corrected. The hand-off package table now lists the
five files added since v1.7.9, and its open-items section is scoped to judgements
rather than unverified claims.

## Also expected

| Check | Expected |
|---|---|
| `Sources!D36` | still blank |
| `Sources!E36` | still `open`, a placeholder citing nothing |
| sharedStrings `<si>` count vs `uniqueCount` | equal |
| `CHECK` markers left in `Sources` | three, all inside sentences recording a resolution |
| `Open Items!AA66` | rewritten; states the divergence bound as 0.78, not 0.76 |
