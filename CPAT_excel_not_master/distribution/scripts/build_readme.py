#!/usr/bin/env python3
"""
Build cpat_excel/readme.xlsx: a table of what the pipeline does, generated
from common.MODULES and the tiers actually on disk (so it can't drift out of
sync with the pipeline the way a hand-written README could).

Usage:
  python build_readme.py
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

import common

OUT_PATH = os.path.join(common.CPAT_EXCEL_DIR, "readme.xlsx")

TITLE_FONT = Font(name="Arial", size=14, bold=True)
HEADER_FONT = Font(name="Arial", size=10, bold=True, color="FFFFFF")
BASE_FONT = Font(name="Arial", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
from openpyxl.styles import PatternFill
HEADER_FILL = PatternFill("solid", fgColor="4472C4")

FOLDER_ROWS = [
    ("original/", "source", "-", "Untouched source workbook: the original CPAT climate policy "
     "model (.xlsb), one tab per module/data source.", "-"),
    ("<Module>/data_bymodule/", "1 - by module", "forward", "Raw tab-for-tab copy of that module's "
     "source sheets, values only, at their original row/column positions.", "build_from_source.py"),
    ("<Module>/data_tabular/", "2 - tabular", "forward", "One clean, single-header table per logical "
     "table: sheets that bundled several tables are split apart, merged-cell blanks are "
     "forward-filled, duplicate rows are dropped.", "build_from_source.py"),
    ("<Module>/data_standardized/", "3 - standardized", "forward", "Long/tidy form of every "
     "measure-bearing table (id columns + variable + value), so tables can be stacked/combined "
     "across modules. Crosswalk/lookup tables are already tidy and are carried over unchanged.",
     "build_from_source.py"),
    ("<Module>/data_tabular_regenerated/", "2' - tabular (checked)", "reverse", "Rebuilt by "
     "pivoting data_standardized back to wide form. Diffed against data_tabular to confirm the "
     "standardized tier lost nothing the tabular tier needs.", "tabular_from_standardized.py"),
    ("<Module>/data_bymodule_regenerated/", "1' - by module (checked)", "reverse", "Rebuilt by "
     "regrouping data_tabular_regenerated's tables back onto their source-sheet tabs. Simple "
     "single-table sheets are diffed exactly against data_bymodule; sheets that bundled several "
     "tables (e.g. Mapping) are rebuilt as a readable stack of title/header/data blocks rather "
     "than a byte-identical replica of the original layout.", "bymodule_from_tabular.py"),
]


def build_folders_sheet(wb):
    ws = wb.create_sheet("Pipeline")
    ws["A1"] = "cpat_excel data pipeline"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (
        "Forward:  original.xlsb -> data_bymodule -> data_tabular -> data_standardized\n"
        "Reverse:  data_standardized -> data_tabular_regenerated -> data_bymodule_regenerated "
        "(round-trip check that nothing was lost)"
    )
    ws["A2"].font = BASE_FONT
    ws["A2"].alignment = WRAP
    ws.merge_cells("A2:E2")
    ws.row_dimensions[2].height = 32

    headers = ["Folder", "Tier", "Direction", "Description", "Produced by"]
    header_row = 4
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for r, row in enumerate(FOLDER_ROWS, start=header_row + 1):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = BASE_FONT
            cell.alignment = WRAP
        ws.row_dimensions[r].height = 48

    widths = [30, 20, 12, 80, 28]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = f"A{header_row}:E{header_row + len(FOLDER_ROWS)}"
    return ws


def build_modules_sheet(wb):
    ws = wb.create_sheet("Modules")
    headers = ["Module", "Output file", "Source sheets (original.xlsb)", "# tables (data_tabular)"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    r = 2
    for module_name, spec in common.MODULES.items():
        tabular_path = os.path.join(
            common.module_dir(module_name, "data_tabular"), spec["output_name"]
        )
        n_tables = len(common.read_tables_xlsx(tabular_path)) if os.path.exists(tabular_path) else 0
        row = [
            module_name,
            spec["output_name"],
            ", ".join(spec["source_sheets"]),
            n_tables,
        ]
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = BASE_FONT
            cell.alignment = WRAP
        ws.row_dimensions[r].height = 32
        r += 1

    widths = [16, 32, 55, 20]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:D{r - 1}"
    return ws


def build_tables_sheet(wb):
    ws = wb.create_sheet("Tables")
    headers = [
        "Module", "Table", "Source sheet", "Tabular (rows x cols)",
        "Standardized (rows x cols)", "Long-formatted?", "Standardized id columns",
    ]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    r = 2
    for module_name, spec in common.MODULES.items():
        tabular_path = os.path.join(
            common.module_dir(module_name, "data_tabular"), spec["output_name"]
        )
        standard_path = os.path.join(
            common.module_dir(module_name, "data_standardized"), spec["output_name"]
        )
        if not (os.path.exists(tabular_path) and os.path.exists(standard_path)):
            continue
        tabular_tables = common.read_tables_xlsx(tabular_path)
        standard_tables = common.read_tables_xlsx(standard_path)

        for table_name, (th, td) in tabular_tables.items():
            sh, sd = standard_tables.get(table_name, ([], []))
            melt_spec = spec["melt_specs"].get(table_name)
            row = [
                module_name,
                table_name,
                spec["table_to_sheet"].get(table_name, ""),
                f"{len(td)} x {len(th)}",
                f"{len(sd)} x {len(sh)}",
                "Yes" if melt_spec else "No (already tidy)",
                ", ".join(melt_spec["id_cols"]) if melt_spec else "",
            ]
            for c, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.font = BASE_FONT
                cell.alignment = WRAP
            r += 1

    widths = [14, 30, 16, 20, 24, 20, 50]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:G{r - 1}"
    return ws


def main():
    wb = Workbook()
    wb.remove(wb.active)
    build_folders_sheet(wb)
    build_modules_sheet(wb)
    build_tables_sheet(wb)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")


if __name__ == "__main__":
    main()
