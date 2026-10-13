# Standalone Distribution module (Egypt)

`CPAT_Distribution_Standalone_Egypt.xlsx` is a standalone rebuild of the household
distributional/incidence module from `cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb`
(sheet "Distribution"), scoped to Egypt only. See the workbook's own **ReadMe** tab for
scope, structure and verification notes.

## Pipeline

```
extract_egypt.py   reads the source .xlsb (pyxlsb) + cpat_excel/Distribution/data_tabular,
                    filters everything to Egypt, and writes egypt_data.pkl

build_workbook.py   reads egypt_data.pkl and writes:
                      - CPAT_Distribution_Standalone_Egypt.xlsx (data tabs, Distribution_
                        Inputs, Distribution_Outputs, ReadMe -- all labels/structure/
                        styling, but with formula cells and the LAMBDA library left BLANK)
                      - RebuildDistributionModule.bas (a VBA macro that adds the LAMBDA
                        names and writes every formula cell via Excel's own object model)

reference_calc.py   a pure-Python re-implementation of the same LAMBDA logic, used to
                    validate the design against the source workbook's cached Egypt values
                    before it was encoded as Excel formulas. Not part of the shipped
                    workbook; kept for anyone who wants to re-verify or extend the model.
```

### Why the .bas file

Excel's .xlsx format requires every "future function" (LAMBDA, LET, XLOOKUP, HSTACK,
SCAN, MAKEARRAY -- anything added after the OOXML spec froze) to be written internally
with an `_xlfn.` prefix, and every LAMBDA/LET parameter name to carry an `_xlpm.` prefix,
in both its declaration and every reference to it. Getting this byte-for-byte right by
hand-writing XML (via openpyxl) proved unreliable across Excel versions/builds -- Excel
would silently strip the offending names/formulas and show a "needs repair" prompt.
Shipping the workbook with everything left blank, plus a macro that recreates the LAMBDA
names and formulas through Excel's own `Names.Add` / `Range.Formula2` API, sidesteps the
encoding problem entirely: Excel does its own internal encoding, guaranteed correct.

To use: open the .xlsx, Alt+F11, Insert > Module, paste in `RebuildDistributionModule.bas`,
then run `RebuildDistributionModule` (F5). It defines all 39 named LAMBDA functions plus
the plain named ranges, then writes all ~1,400 formula cells in Distribution_Inputs and
Distribution_Outputs, then forces a full recalculation.

Re-run with:

```
python3 extract_egypt.py     # needs cpat_excel/original/CPAT 1.0pre_456_NoPropData.xlsb
python3 build_workbook.py    # needs egypt_data.pkl
```

`egypt_data.pkl` is checked in so `build_workbook.py` can be re-run (e.g. to tweak
formatting or add sections) without re-reading the large source `.xlsb`.
