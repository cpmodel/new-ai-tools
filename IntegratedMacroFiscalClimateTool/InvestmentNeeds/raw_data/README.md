Optional inputs for `build/build_investment_model.py`.

- `ctryprem.xlsx` — Damodaran, *Country Default Spreads and Risk Premiums*
  (https://pages.stern.nyu.edu/~adamodar/ → Data → Current data → country risk).
  If this file is present, the build script loads the sheet that has `Country` and
  `Default Spread` headers into the workbook's Damodaran sheet. Check the `Matched?`
  column after the build: names that don't match a country are left blank.
