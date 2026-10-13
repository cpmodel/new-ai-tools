"""Replace the Egypt rows of data/prices_dom.csv with the corrected Egypt price block supplied by the user.

Input:  data/source/Egypt_Price_Data_2026-10-09.xlsx (one sheet, legacy Prices_dom layout: bands rows 1-4, codes
        row 5, one row per year 2019-2024; extra info columns spsrc/rprel/rpsrc/taxa/last are not used).
Output: data/prices_dom.csv (EGY rows for the columns the file shares with the dataset replaced; years missing
        from the dataset appended), data/prices_dom_changes.csv (every changed cell: country_year, column, old,
        new) so the workbook can mark them.

    python update_prices_egypt_v0_1.py
"""
import csv
import os

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, 'data', 'source', 'Egypt_Price_Data_2026-10-09.xlsx')
DOM = os.path.join(HERE, 'data', 'prices_dom.csv')
CHANGES = os.path.join(HERE, 'data', 'prices_dom_changes.csv')
ID_COLS = ('country_year', 'countrycode_weo', 'countrycode', 'countryname', 'year')


def fmt(v):
    if v is None:
        return ''
    if isinstance(v, float):
        return repr(round(v, 9)).rstrip('0').rstrip('.') if v != int(v) else str(int(v))
    return str(v)


def same(a, b):
    try:
        return abs(float(a or 0) - float(b or 0)) < 1e-9
    except ValueError:
        return a == b


def main():
    rows = list(openpyxl.load_workbook(SRC, data_only=True).active.iter_rows(values_only=True))
    hdr = rows[4]
    new = {r[0]: dict(zip(hdr, r)) for r in rows[5:] if r[0]}
    with open(DOM, encoding='utf-8') as f:
        data = list(csv.DictReader(f))
    cols = list(data[0].keys())
    shared = [c for c in cols if c in hdr and c not in ID_COLS]
    changes, seen = [], set()
    for row in data:
        cy = row['country_year']
        if cy in new:
            seen.add(cy)
            for c in shared:
                v = fmt(new[cy][c])
                if not same(row[c], v):
                    changes.append((cy, c, row[c], v))
                row[c] = v
    added = []
    for cy in sorted(set(new) - seen):            # years not in the dataset: inserted before the country's rows
        row = {c: fmt(new[cy].get(c)) for c in cols}
        added.append(row)
        changes += [(cy, c, '', row[c]) for c in shared if row[c] != '']
    if added:
        i = next(k for k, r in enumerate(data) if r['countrycode'] == added[0]['countrycode'])
        data[i:i] = added                         # original row order otherwise kept
    with open(DOM, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=cols, lineterminator='\r\n')   # as the extracted file
        w.writeheader()
        w.writerows(data)
    with open(CHANGES, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, lineterminator='\n')
        w.writerow(['country_year', 'column', 'old', 'new', 'source'])
        w.writerows([c + ('Egypt_Price_Data_2026-10-09.xlsx (user, corrected Egypt price block)',) for c in changes])
    print(f'{len(changes)} cells changed or added ({len(set(new) - seen)} new years); '
          f'{len(set(hdr) - set(cols) - {None})} source columns not in the dataset (info only)')


if __name__ == '__main__':
    main()
