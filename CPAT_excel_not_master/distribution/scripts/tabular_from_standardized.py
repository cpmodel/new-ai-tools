#!/usr/bin/env python3
"""
Reverse pipeline, step 1: data_standardized -> data_tabular_regenerated

Reads each module's long/tidy workbook and pivots every melted table back to
its wide tabular form (tables that were already tidy, e.g. crosswalks, are
carried over unchanged). Writes data_tabular_regenerated/<module>.xlsx and
diffs it against the original data_tabular/<module>.xlsx to confirm the
standardized tier lost nothing the tabular tier needs.

Usage:
  python tabular_from_standardized.py [module_name ...]   # default: all modules
"""
import os
import sys

import common


def regenerate_module(module_name):
    spec = common.MODULES[module_name]
    print(f"=== {module_name} ===")

    standard_path = f"{common.module_dir(module_name, 'data_standardized')}/{spec['output_name']}"
    standard_tables = common.read_tables_xlsx(standard_path)

    tabular_tables = {}
    for name, (header, data) in standard_tables.items():
        melt_spec = spec["melt_specs"].get(name)
        if melt_spec is None:
            tabular_tables[name] = (header, data)
            continue
        out_header, out_data = common.pivot_rows(
            header, data, melt_spec["id_cols"], melt_spec["var_name"], melt_spec["value_name"]
        )
        tabular_tables[name] = (out_header, out_data)

    out_path = f"{common.module_dir(module_name, 'data_tabular_regenerated')}/{spec['output_name']}"
    common.write_tables_xlsx(out_path, tabular_tables)
    print(f"Wrote {out_path}")
    for name, (h, d) in tabular_tables.items():
        print(f"  {name}: {len(d)} rows x {len(h)} cols")

    original_path = f"{common.module_dir(module_name, 'data_tabular')}/{spec['output_name']}"
    if os.path.exists(original_path):
        original_tables = common.read_tables_xlsx(original_path)
        print(f"Comparing against {original_path}:")
        common.compare_tables(original_tables, tabular_tables, label=module_name)
    print()


def main():
    modules = sys.argv[1:] or list(common.MODULES.keys())
    for m in modules:
        if m not in common.MODULES:
            raise SystemExit(f"Unknown module '{m}'. Known modules: {list(common.MODULES)}")
        regenerate_module(m)


if __name__ == "__main__":
    main()
