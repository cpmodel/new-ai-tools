#!/usr/bin/env python3
"""
Forward pipeline: original/*.xlsb -> data_bymodule -> data_tabular -> data_standardized

For each module registered in common.MODULES, reads the listed source sheets
out of the original CPAT .xlsb workbook and writes three increasingly clean
xlsx workbooks:

  data_bymodule/<module>.xlsx      raw tab-for-tab copy of the source sheets
  data_tabular/<module>.xlsx       one clean single-header table per logical table
  data_standardized/<module>.xlsx  long/tidy form of the measure-bearing tables

Usage:
  python build_from_source.py [module_name ...]   # default: all modules
"""
import sys

from pyxlsb import open_workbook

import common


def load_grid(sheet_name):
    with open_workbook(common.SOURCE_XLSB) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            rows = list(sheet.rows())
    ncols = max((len(r) for r in rows), default=0)
    grid = []
    for r in rows:
        rowvals = [None] * ncols
        for c in r:
            rowvals[c.c] = common.clean_value(c.v)
        grid.append(rowvals)
    return grid


def build_module(module_name):
    spec = common.MODULES[module_name]
    print(f"=== {module_name} ===")

    print("Loading source sheets from xlsb...")
    grids = {name: load_grid(name) for name in spec["source_sheets"]}
    for name, g in grids.items():
        ncols = max((len(r) for r in g), default=0)
        print(f"  {name}: {len(g)} rows x {ncols} cols")

    # --- Stage 1: data_bymodule -------------------------------------------------
    bymodule_path = f"{common.module_dir(module_name, 'data_bymodule')}/{spec['output_name']}"
    common.write_raw_grids_xlsx(bymodule_path, grids)
    print(f"Wrote {bymodule_path}")

    # --- Stage 2: data_tabular ---------------------------------------------------
    tabular_tables = {}

    for name in spec["simple_sheets"]:
        g = grids[name]
        header, data = common.subgrid(g, 0, 1, len(g) - 1, list(range(len(g[0]))))
        tabular_tables[name] = (header, data)

    for sheet_name, table_specs in spec["composite_sheets"].items():
        g = grids[sheet_name]
        for t in table_specs:
            header, data = common.subgrid(
                g, t["header_row"], t["data_start"], t["data_end"], t["cols"]
            )
            if t["ffill"]:
                header, data = common.ffill_cols(header, data, t["ffill"])
            if t["dedupe"]:
                header, data = common.dedupe_rows(header, data)
            tabular_tables[t["table"]] = (header, data)

    tabular_path = f"{common.module_dir(module_name, 'data_tabular')}/{spec['output_name']}"
    common.write_tables_xlsx(tabular_path, tabular_tables)
    print(f"Wrote {tabular_path}")
    for name, (h, d) in tabular_tables.items():
        print(f"  {name}: {len(d)} rows x {len(h)} cols")

    # --- Stage 3: data_standardized -----------------------------------------------
    standard_tables = {}
    for name, (header, data) in tabular_tables.items():
        melt_spec = spec["melt_specs"].get(name)
        if melt_spec is None:
            standard_tables[name] = (list(header), [list(r) for r in data])
            continue
        out_header, out_data = common.melt_rows(
            header, [list(r) for r in data],
            melt_spec["id_cols"], melt_spec["var_name"], melt_spec["value_name"],
        )
        if melt_spec["var_name"] == "year":
            # tidy 1960.0-style floats from the source header into plain ints
            yi = out_header.index("year")
            for row in out_data:
                if isinstance(row[yi], float):
                    row[yi] = int(row[yi])
        standard_tables[name] = (out_header, out_data)

    standard_path = f"{common.module_dir(module_name, 'data_standardized')}/{spec['output_name']}"
    common.write_tables_xlsx(standard_path, standard_tables)
    print(f"Wrote {standard_path}")
    for name, (h, d) in standard_tables.items():
        print(f"  {name}: {len(d)} rows x {len(h)} cols")
    print()


def main():
    modules = sys.argv[1:] or list(common.MODULES.keys())
    for m in modules:
        if m not in common.MODULES:
            raise SystemExit(f"Unknown module '{m}'. Known modules: {list(common.MODULES)}")
        build_module(m)


if __name__ == "__main__":
    main()
