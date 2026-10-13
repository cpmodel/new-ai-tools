# Lessons learned: building a LAMBDA-based Excel workbook safely

This document exists because almost every bug encountered while building this
workbook compiled, ran, and showed no error the first time it was tried — each
one only surfaced under a specific condition (a particular Excel build, a
specific table name, calling a formula with an array instead of a scalar).
None of them were caught by "does it open without an error" or "does a quick
scalar test give the right number." The rules below are written so the next
person (or the next LLM) building something like this doesn't have to
rediscover each one by hand in real Excel.

Each section: what the bug looked like, why it happened, and the rule to
follow instead.

## 1. Never hand-write `_xlfn`/`_xlpm`-encoded formulas into the XML

**Symptom:** "Excel completed file level validation and repair." Removed
Records: Named range, Formula. `#NAME?` on every LAMBDA call. No LAMBDA
functions visible in the Name Manager at all.

**Cause:** Every Excel "future function" (`LAMBDA`, `LET`, `XLOOKUP`,
`HSTACK`, `SCAN`, `MAKEARRAY`, ...) must be written in the underlying XML
with an `_xlfn.` prefix (`_xlfn._xlws.` for some newer worksheet-function-group
ones), and every LAMBDA/LET parameter name needs an `_xlpm.` prefix on its
declaration *and every reference to it*. Getting this byte-for-byte right by
constructing the XML directly (e.g. via openpyxl) is unreliable across Excel
builds — small encoding mistakes get silently stripped by Excel's repair
pass instead of raising a clear error.

**Rule:** Don't fight this encoding by hand. Ship the `.xlsx` with every
formula cell and every defined name left blank (so it opens with zero
repair prompts), and populate everything through a VBA macro using
`Names.Add` / `Range.Formula2`. Excel performs its own internal `_xlfn`/
`_xlpm` encoding when you set a formula through its own object model, so
this class of bug cannot happen. This is the single biggest architectural
decision in this project and the one that actually ended the "repair
needed" bug reports.

## 2. Defined names and Excel Table names share ONE case-insensitive namespace

**Symptom:** A LAMBDA function compiles, the macro runs with no error, but
calling it gives `#REF!` everywhere, and the Name Manager shows "only an
`ElasAdj` Range or Table" instead of the expected LAMBDA name. The macro
reported success.

**Cause:** `ELASTADJ` (the LAMBDA name) and `ElastAdj` (an existing Excel
Table's `displayName`) collided — Excel's name resolution is
case-insensitive across *both* defined names and table names. `Names.Add`
for `ELASTADJ` failed, but the macro's `On Error Resume Next` (needed so one
missing/duplicate name doesn't abort the whole batch) swallowed the error
silently.

**Rule:**
- Never assume `On Error Resume Next` is safe without a separate way to
  detect what it swallowed.
- At build time, before generating the VBA, check every defined name you're
  about to create against every Excel Table `displayName` you're about to
  create, case-insensitively. Fail the build (don't just warn) on any
  collision. This project's `build_workbook.py` does this permanently now
  (see the "Checked N names against M table names" build step) — if you
  copy this pattern elsewhere, keep that check.

## 3. `MIN`/`MAX` do not broadcast over array arguments — `IF` does

**Symptom:** A formula that calls a LAMBDA with a scalar decile works fine.
The exact same LAMBDA, called with the full array of all 10 deciles at once,
returns `#VALUE!` or a single collapsed number instead of 10 values.

**Cause:** `MIN(0, {array of 10 values})` and `MAX(...)` do **not** clamp
each array element independently — Excel reduces *all* arguments (scalars
and array elements alike) down to one single value. This looks exactly like
an element-wise clamp when you test it with a scalar second argument, which
is why it passed every test until it was called with `DECILE_ARRAY`.

**Rule:** Never use `MIN`/`MAX` to clamp one argument against another when
either argument can be array-shaped. Use `IF(condition, a, b)` instead —
`IF` broadcasts correctly element-wise over arrays. `MIN`/`MAX` are fine
only when you're certain every argument is a true scalar (e.g. two
workbook-level totals).

## 4. `INDEX(range, 0, X)` is only safe when the *other* index is scalar

**Symptom:** Same shape as #3 — works with a scalar decile, `#VALUE!` with
the full `DECILE_ARRAY`.

**Cause:** `INDEX(range, 0, X)` ("0 = whole column/row") is well-defined
when the index you *do* supply is a scalar. It is **not** a supported way to
pull several columns/rows at once by passing an array for `X` while the
other index is `0`.

**Rule:** To select a row/column by a scalar key and then pick out a
possibly array-valued sub-selection from it: first `MATCH` to find the row
(always scalar, since the key being matched is always scalar) →
`INDEX(range, row, 0)` to pull that whole row (scalar row index, so this
direction is well-defined) → `CHOOSECOLS(row, X)` to select the (possibly
array-valued) columns you actually want. `CHOOSECOLS` is the function
actually designed to take an array of column positions; `INDEX`'s `0`
shorthand is not.

## 5. Scalar-only tests do not catch array-broadcasting bugs — test with the real array

Both #3 and #4 compiled, ran, and gave correct answers under manual
spot-checking with a single decile. They only failed once called with the
full `DECILE_ARRAY` — which is also the call pattern actually used
throughout the real output sheet. A smoke test that only ever tries scalar
inputs will not find this class of bug.

**Rule:** Any function in a library like this one that is ever called with
an array argument needs at least one test that calls it with that *same*
array shape, not just a scalar stand-in. This project's `Tests` sheet
deliberately calls every LAMBDA with the full `DECILE_ARRAY` for exactly
this reason — see `build_workbook.py`'s `TEST_CASES` list.

## 6. Add a static check for the *shape* of a bug, not just a regression test for the one instance

After fixing #3 and #4, `build_workbook.py` gained a build-time scan that
parses every LAMBDA formula's text and flags:
- any `INDEX(range, 0, X)` where `X` textually references `decile` (the
  parameter name that carries array values throughout this codebase), and
- any `MIN`/`MAX` call with more than one argument where an argument
  references `decile`.

This catches a *new* instance of either bug shape anywhere in the LAMBDA
library, immediately at build time, before it ever reaches Excel — not just
the one specific formula that broke before. When you fix a bug caused by a
general pattern, prefer writing a check for the pattern over a regression
test for the one broken formula.

## 7. Chunk long strings into VBA line-continuations BEFORE escaping, not after

**Symptom:** A LAMBDA's formula works for some decile values and is subtly
wrong for others, with no error at all — the macro ran cleanly and every
line was valid VBA.

**Cause:** When generating a VBA string literal for a long formula, the
escaping step (`"` → `""`) was applied to the *whole* string first, and
*then* the result was sliced into fixed-width chunks for line-continuation.
If a chunk boundary happened to fall between the two characters of a
doubled `""` escape pair, it silently split one escape sequence into two —
both halves remained individually valid VBA syntax (so nothing failed to
compile), but the string decoded back to something subtly different from
what was intended.

**Rule:** Chunk the **raw** text first, then escape each chunk
independently. A chunk boundary can then only ever fall between two
original characters, never inside the two-character encoding of one — this
makes the bug structurally impossible rather than just unlikely.

## 8. Verify generated code round-trips, with a real parser — not string-replace

Once chunk-then-escape is in place, add an automatic check that decodes the
generated VBA string-concatenation expression back into the original text
and asserts they're equal, for every single generated statement. Write the
decoder as a real character-by-character parser, not a string-replace or
regex — a naive approach can be fooled by legitimate characters inside the
literal (for example, this codebase's `HHKEY` formula contains a real `&`
string-concatenation operator, which a `.replace("&", ...)`-based decoder
would misinterpret as a VBA-level string-join operator). This project's
`_parse_vba_string_concat` / `_verify_vba_string_roundtrip` run automatically
on every build, against all ~1,460 generated statements.

## 9. Excel Tables need unique column headers — duplicates break silently (sort of)

**Symptom:** "Excel completed file level validation and repair." Repaired
Records: Table.

**Cause:** One of the source data tables
(`Mapping_CPATSectorsToISIC`) has the literal column header `"CPAT code"`
twice (once for a 5-category code, once for a 15-category code) — legitimate
in a loose CSV/export sense, but an Excel Table (`ListObject`) requires
unique column names and Excel repairs the table on open rather than
rejecting the file outright.

**Rule:** When turning tabular data into an Excel Table, de-duplicate
column headers at build time (e.g. append a numeric suffix to the second
and later occurrence of a repeated name) rather than trusting the source
data to already have unique headers.

## 10. Validate formula *design* in Python before encoding it as an Excel formula

Every LAMBDA in this workbook was first written and checked as a plain
Python function (`reference_calc.py`) against the original workbook's own
cached values, *before* being translated into an Excel formula string. This
caught real derivation bugs early (wrong price-change layer, a missing
elasticity-adjustment multiplier, a wrong PIT baseline normalization) at the
cheapest possible point to fix them — before any Excel/VBA-specific
encoding issues could obscure whether a wrong number was a formula-design
bug or an Excel-mechanics bug.

This did **not**, by itself, catch bugs #2–#4 above — those are specific to
how Excel's calculation engine actually behaves (case-insensitive names,
`MIN`/`MAX`/`INDEX` broadcasting rules) and only show up once you run the
real formula in real Excel. Python validation and in-Excel testing are
complementary, not substitutes for each other: use Python to validate that
the *design* produces the right numbers, and real Excel (or, better, the
`Tests` sheet calling live Excel formulas) to validate that the *encoding*
of that design behaves the way you assumed.

## Summary checklist for the next LAMBDA-heavy workbook

- [ ] Ship with blank formulas/names; populate via a VBA macro through
      `Names.Add`/`Range.Formula2` — never hand-encode `_xlfn`/`_xlpm` XML.
- [ ] Build-time check: every defined name vs. every Excel Table name,
      case-insensitively, fail on collision.
- [ ] Never `MIN`/`MAX` an argument that could be array-shaped — use `IF`.
- [ ] Never `INDEX(range, 0, X)` where `X` could be array-shaped — use
      `MATCH` + `INDEX(range, row, 0)` + `CHOOSECOLS`.
- [ ] Test every array-callable function with its real array shape, not
      just a scalar stand-in.
- [ ] Add a static scan for each bug *pattern* you fix, not just a
      regression test for the one formula.
- [ ] Chunk long generated strings before escaping, not after; verify the
      round-trip with a real parser.
- [ ] De-duplicate Excel Table column headers at build time.
- [ ] Validate formula design in a plain-language (e.g. Python) model
      before encoding it as an Excel formula, in addition to — not instead
      of — testing the real encoded formula in real Excel.
