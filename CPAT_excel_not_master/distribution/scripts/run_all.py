#!/usr/bin/env python3
"""
Run the full cpat_excel pipeline end to end for every registered module:

  forward:  original.xlsb -> data_bymodule -> data_tabular -> data_standardized
  reverse:  data_standardized -> data_tabular_regenerated -> data_bymodule_regenerated

The reverse pass exists to check that data_standardized retains what the
tabular and bymodule tiers need; each reverse script prints a diff against
the corresponding forward-pass output.

Usage:
  python run_all.py [module_name ...]   # default: all modules in common.MODULES
"""
import sys

import bymodule_from_tabular
import build_from_source
import common
import tabular_from_standardized


def main():
    modules = sys.argv[1:] or list(common.MODULES.keys())
    for m in modules:
        if m not in common.MODULES:
            raise SystemExit(f"Unknown module '{m}'. Known modules: {list(common.MODULES)}")

    print("########## FORWARD: source -> bymodule -> tabular -> standardized ##########")
    for m in modules:
        build_from_source.build_module(m)

    print("########## REVERSE: standardized -> tabular_regenerated ##########")
    for m in modules:
        tabular_from_standardized.regenerate_module(m)

    print("########## REVERSE: tabular_regenerated -> bymodule_regenerated ##########")
    for m in modules:
        bymodule_from_tabular.regenerate_module(m, "data_tabular_regenerated")


if __name__ == "__main__":
    main()
