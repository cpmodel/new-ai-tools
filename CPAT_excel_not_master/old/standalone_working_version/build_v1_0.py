"""Build CPAT_Industry_Kernel_Egypt_v1.0.xlsx (final release) from v0.17: relabel only, no formula changes.

Run with Excel installed (from this folder):  python build_v1_0.py
"""
import os
import shutil

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.17.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v1.0.xlsx")


def main():
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(DST)
        ws = wb.Worksheets("Settings")
        assert ws.Range("A44").Value == "v0.17"
        ws.Range("A1").Value = "CPAT industry kernel - Egypt, final v1.0"
        ws.Range("A44").EntireRow.Copy()
        ws.Range("A45").PasteSpecial(-4122)  # formats
        ws.Range("A45").Value = "v1.0"
        ws.Range("B45").Value = "2026-10-04"
        ws.Range("C45").Value = ("Final release: v0.17 relabelled, no formula or value changes. Final Table 2 = sheet "
                                 "CarveOut_Table2 (CBAM carve-out). 2030 K: 1A -37.4, 2A -33.1, 2B -34.8, 3A -22.8, "
                                 "3B -19.1, 3C -33.0.")
        xl.CutCopyMode = False
        wb.Save()
        wb.Close(False)
    finally:
        xl.Quit()
    print("saved", DST)


if __name__ == "__main__":
    main()
