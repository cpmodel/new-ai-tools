#!/usr/bin/env python3
"""
Reverse pipeline, step 2: data_tabular_regenerated -> data_bymodule_regenerated

Reads each module's clean tabular workbook and regroups its tables back onto
the source-sheet tabs they came from (common.MODULES[...]['table_to_sheet']).
A sheet that bundled several logical tables (e.g. Mapping) is rebuilt as a
sequence of title/header/data blocks, in the sheet's original top-to-bottom
order; sheets that were already a single table are carried over unchanged.

This does not attempt a byte-identical reconstruction of the original
layout: data_tabular's cleanup (forward-filling merged cells, dropping
duplicate rows) is intentionally irreversible, and the original's quirky
side-by-side column layout (several tables sharing one row range) isn't
recoverable once the tables have been split apart. What's rebuilt is a
readable, unambiguous stack of the same tables, structurally where they
started.

Usage:
  python bymodule_from_tabular.py [module_name ...]   # default: all modules

  By default reads from data_tabular_regenerated. Pass --from-original to
  read data_tabular instead (useful to test this step in isolation).
"""
import os
import sys

import common


def build_composite_grid(table_specs, tabular_tables):
    grid = []
    for spec in table_specs:
        name = spec["table"]
        header, data = tabular_tables[name]
        title = spec["title"] or name.replace("_", " ")
        grid.append([title])
        grid.append(list(header))
        for row in data:
            grid.append(list(row))
        grid.append([])
        grid.append([])
    return grid


def regenerate_module(module_name, tabular_tier="data_tabular_regenerated"):
    spec = common.MODULES[module_name]
    print(f"=== {module_name} (from {tabular_tier}) ===")

    tabular_path = f"{common.module_dir(module_name, tabular_tier)}/{spec['output_name']}"
    tabular_tables = common.read_tables_xlsx(tabular_path)

    grids = {}

    # simple sheets: one table -> one sheet, unchanged
    simple_table_to_sheet = {
        t: s for t, s in spec["table_to_sheet"].items() if s in spec["simple_sheets"]
    }
    for table_name, sheet_name in simple_table_to_sheet.items():
        header, data = tabular_tables[table_name]
        grids[sheet_name] = [list(header)] + [list(r) for r in data]

    # composite sheets: several tables stacked back onto one tab
    for sheet_name, table_specs in spec["composite_sheets"].items():
        grids[sheet_name] = build_composite_grid(table_specs, tabular_tables)

    out_path = f"{common.module_dir(module_name, 'data_bymodule_regenerated')}/{spec['output_name']}"
    common.write_raw_grids_xlsx(out_path, grids)
    print(f"Wrote {out_path}")
    for name, g in grids.items():
        ncols = max((len(r) for r in g), default=0)
        print(f"  {name}: {len(g)} rows x {ncols} cols")

    original_path = f"{common.module_dir(module_name, 'data_bymodule')}/{spec['output_name']}"
    if os.path.exists(original_path):
        original_grids_wb = common.read_tables_xlsx(original_path)
        print(f"Comparing simple sheets against {original_path}:")
        expected = {}
        actual = {}
        for table_name, sheet_name in simple_table_to_sheet.items():
            oh, od = original_grids_wb[sheet_name]
            expected[sheet_name] = (list(oh), [list(r) for r in od])
            gh = grids[sheet_name][0]
            gd = grids[sheet_name][1:]
            actual[sheet_name] = (gh, gd)
        common.compare_tables(expected, actual, label=module_name)
        composite_names = ", ".join(spec["composite_sheets"].keys())
        print(f"  [{module_name}] {composite_names}: intentionally restructured, not compared")
    print()


def main():
    args = sys.argv[1:]
    from_original = "--from-original" in args
    args = [a for a in args if a != "--from-original"]
    tabular_tier = "data_tabular" if from_original else "data_tabular_regenerated"

    modules = args or list(common.MODULES.keys())
    for m in modules:
        if m not in common.MODULES:
            raise SystemExit(f"Unknown module '{m}'. Known modules: {list(common.MODULES)}")
        regenerate_module(m, tabular_tier)


if __name__ == "__main__":
    main()
