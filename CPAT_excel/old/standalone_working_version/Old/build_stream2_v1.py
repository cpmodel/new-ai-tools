"""Build CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v1.xlsx from v0.7 (Task F: energy CO2 in the kernel).

Stream 2 branch off v0.7 (Stream 1 continues on v0.8+ in the same checkout); to be merged into the main line later.
Only NEW sheets / appended rows are added, in positions Stream 1 does not use, so the merge is additive:

  Data_EF             New sheet (after Data_Elast). Fuel CO2 emission factors for industry, Egypt, exactly as CPAT
                      uses them for energy-related CO2 (CPAT_1_0pre_456, Mitigation):
                        EF used (tCO2/ktoe) = IIASA EF (tCO2/GJ, rows 1001-1009, sector 'ind', switch
                        EmissionsFactCO2 = IIASA*) x 41868 GJ/ktoe (row 1035) x inventory adjustment factor
                        0.9296150140376861 (row 977; applied because Egypt is Non-Annex I and MTInputs
                        EFsAdjustmentNonAnnexI = Yes*).  Reproduces CPAT column B of rows 6516-6552 to the last bit.
                      Codes egy.mit.efc.<fuel> (column A), value read by the kernel in column G.
                      bio = 0 (biogenic, CPAT row 1009) and ren = 0 (no combustion), as in CPAT.
  Emissions_Industry  New sheet (after Mitigation_Industry) mirroring its layout (columns A-AI, years K..AI, data
                      L..AE = 2022-2041). Per scenario (1 baseline, 2 policy) and sector (irn, cem, nfm, mch):
                      9 fuel rows  eco2 = Mitigation_Industry ener (ktoe) x Data_EF EF (tCO2/ktoe) / 1e6, MtCO2
                      (same unit as the CBAM block emisf/emisp rows), output codes
                      egy.mit.eco2.<sector>.<fuel>.e.<s>, plus a total row egy.mit.eco2.<sector>.tot.e.<s>.
                      Why a new sheet: Mitigation_Industry has no spare rows after the sector sections (single blank
                      separator rows 64/124/184/244 and 463/523/583/643, CBAM block follows) and the CBAM-block spare
                      rows are used by Stream 1 (Tasks C/K). Row references into Mitigation_Industry are by cell, so
                      a merge only needs the ener rows to stay where they are (asserted below).
  Check               Section 'Task F (Stream2 v1)' appended at the end: (a) sector total - sum of fuel rows,
                      (b) Data_EF EF used - CPAT EF used, (c) baseline sector energy CO2 vs CPAT egy.mit.co2.ind.<sector>.1
                      (Mitigation rows 6571-6575, mt.CO2, 2022-2041, stored as values), (d) count of fuel-years with
                      scenario-2 CO2 above scenario-1 CO2 while the explicit carbon price (block o1) is positive.
  Settings            Title, C25 note and version-log row 'v0.7branch_Stream2_v1' (appended; the Sheets list rows
                      17-25 is not extended because that would shift rows Stream 1 scripts assert on).
  Scenarios           Note line 'Task F (stream 2, v0.7branch_Stream2_v1): ...'.
  MTInputs            Rows EmissionsFactCO2 and EFsAdjustmentNonAnnexI coloured as used (NORMS 4), nothing else.

Not done here (Task G): replacing the IPPU row or reconciling with product-level fuel emissions.
No existing value changes.

Run with Excel installed (from this folder):  python build_stream2_v1.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLUE, DATA_COLS, FUELS, GREEN, REVIEW, TITLE, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.7.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v1.xlsx")
VER = "v0.7branch_Stream2_v1"
MI, EI = "Mitigation_Industry!", "Emissions_Industry!"

SECTORS = ("irn", "cem", "nfm", "mch")
HEADERS = {1: (5, 65, 125, 185), 2: (404, 464, 524, 584)}   # sector header rows per scenario
ENER0, TOT = 49, 58                                          # ener rows = header + 49 .. +57, total = header + 58
PRICE_ROW = 645                                              # scenario-2 block o1: Carbon Price (Explicit Only)
GJ_PER_KTOE = 41868
ADJ = 0.9296150140376861                                     # CPAT Mitigation row 977 (Egypt)

# IIASA EF tCO2/GJ (CPAT Mitigation rows 1001-1009, sector ind) and CPAT EF used tCO2/ktoe (rows 6516-6552 col B)
EF = {  # fuel: (IIASA tCO2/GJ, CPAT tCO2/ktoe, note)
    "coa": (0.1, 3892.1121407729843, "Mitigation row 1001 (coa, ind)"),
    "nga": (0.055799999999999995, 2171.798574551325, "Mitigation row 1002 (nga, ind)"),
    "gso": (0.067793, 2638.579583594229, "Mitigation row 1003 (gso, ind; 0.0023707 tCO2/l x 28.6 l/GJ)"),
    "die": (0.07184499999999999, 2796.28796753835, "Mitigation row 1004 (die, ind)"),
    "lpg": (0.0686, 2669.9889285702666, "Mitigation row 1005 (lpg, ind)"),
    "ker": (0.067793, 2638.579583594229, "Mitigation row 1006 (ker, ind)"),
    "oop": (0.0767, 2985.2500119728793, "Mitigation row 1007 (oop, ind; labelled 'Crude Oil', 0.4694 tCO2/bbl)"),
    "bio": (0.0, 0.0, "Mitigation row 1009: biomass combustion CO2 counted as 0 (biogenic), as in CPAT"),
    "ren": (0.0, 0.0, "Self-generated renewables: no combustion, EF 0 (CPAT has no EF row)"),
}
FUEL_NAME = {"coa": "Coal", "nga": "Natural gas", "gso": "Gasoline", "die": "Diesel", "lpg": "LPG", "ker": "Kerosene",
             "oop": "Other oil products", "bio": "Biomass", "ren": "Self-generated renewables"}

# CPAT_1_0pre_456 Mitigation rows 6571-6575 (egy.mit.co2.ind.<sector>.1, mt.CO2), columns L..AE = 2022-2041
CPAT_CO2 = {
    "irn": [4.530383518401958, 3.846903361862514, 4.329074620821513, 4.545160654351195, 4.660652723389092,
            4.833079594274428, 5.0529306123252535, 5.266791699961532, 5.469979089775039, 5.677110360427659,
            5.884285846153308, 6.096067414392637, 6.312446688278849, 6.533411878712708, 6.75894772127454,
            6.989035405040347, 7.223652492772068, 7.4627728319721705, 7.706366456313073, 7.9543994769708535],
    "cem": [9.489357446132306, 7.8923978681734726, 8.97404989804158, 9.611738944075816, 9.62559201859446,
            9.86104765555745, 10.224041222352259, 10.57727817828862, 10.913917324515172, 11.25498088660003,
            11.572436051134586, 11.894967219383657, 12.222536887315176, 12.555104231742494, 12.892625020155878,
            13.235051509488896, 13.582332333397817, 13.934412377644659, 14.291232643187334, 14.652730096591432],
    "nfm": [0.0] * 20,
    "mch": [0.0] * 20,
}
CPAT_ROW = {"irn": 6572, "cem": 6575, "nfm": 6573, "mch": 6571}


def check_layout(wb):
    names = [ws.Name for ws in wb.Worksheets]
    for n in ("Data_EF", "Emissions_Industry"):
        if n in names:
            raise ValueError("sheet %s already exists" % n)
    mi = wb.Worksheets("Mitigation_Industry")
    for s, hdrs in HEADERS.items():
        for h, sec in zip(hdrs, SECTORS):
            if mi.Cells(h, 3).Value != sec or mi.Cells(h, 4).Value != "Fuel Type" or mi.Cells(h, 10).Value != s:
                raise ValueError("sector header moved: row %d expected %s scenario %d" % (h, sec, s))
            for i, f in enumerate(FUELS):
                r = h + ENER0 + i
                if mi.Cells(r, 2).Value != f or mi.Cells(r, 6).Value != "ener" or mi.Cells(r, 5).Value != "ktoe":
                    raise ValueError("ener row moved: row %d expected %s" % (r, f))
            if mi.Cells(h + TOT, 2).Value != "tot" or mi.Cells(h + TOT, 6).Value != "ener":
                raise ValueError("ener total row moved: row %d" % (h + TOT))
    if mi.Cells(PRICE_ROW, 3).Value != "Carbon Price (Explicit Only)" or mi.Cells(PRICE_ROW - 1, 10).Value != 2:
        raise ValueError("scenario-2 block price row moved (row %d)" % PRICE_ROW)
    st = wb.Worksheets("Settings")
    if st.Range("B10").Value != "LEGACY":
        raise ValueError("source must be saved with Settings!B10 = LEGACY")
    if st.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")


def mt_param(name):
    return 'INDEX(MTInputs!$F$8:$F$415,MATCH("%s",MTInputs!$H$8:$H$415,0))' % name


def build_data_ef(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Data_Elast"))
    ws.Name = "Data_EF"
    ws.Cells.Font.Name = "Arial"
    for c, w in zip("ABCDEFGHIJK", (22, 24, 6, 12, 12, 12, 12, 14, 10, 22, 90)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("A1").Value = ("Fuel CO2 emission factors for industry, Egypt (Task F, Stream 2) - as used by "
                            "CPAT_1_0pre_456 for energy-related CO2 (Mitigation rows 972-1009, 6511-6552)")
    ws.Range("A1:K1").Interior.Color = TITLE
    ws.Range("A1:K1").Font.Bold = True
    ws.Range("A1:K1").Font.Color = 0xFFFFFF
    ws.Range("A2").Value = ("EF used (tCO2/ktoe) = IIASA EF (tCO2/GJ) x GJ per ktoe x inventory adjustment factor "
                            "(applied when MTInputs EFsAdjustmentNonAnnexI = Yes; Egypt is Non-Annex I). The kernel "
                            "(Emissions_Industry) reads column G by the code in column A.")
    hdr = ("CPAT code", "Fuel", "fuel", "IIASA EF (tCO2/GJ)", "tCO2/ktoe unadjusted", "Adjustment applied",
           "EF used (tCO2/ktoe)", "CPAT EF used (tCO2/ktoe)", "diff G-H", "CPAT EF code", "Source / note")
    ws.Range("A3:K3").Value = hdr
    ws.Range("A3:K3").Interior.Color = GREEN
    ws.Range("A3:K3").Font.Bold = True
    r0 = 4
    for i, f in enumerate(FUELS):
        r = r0 + i
        iiasa, cpat, note = EF[f]
        ws.Range("A%d" % r).Formula = '=Settings!$B$3&".mit.efc."&C%d' % r
        ws.Range("A%d" % r).Interior.Color = BLUE
        ws.Range("B%d:C%d" % (r, r)).Value = (FUEL_NAME[f], f)
        ws.Range("D%d" % r).Value = iiasa
        ws.Range("D%d" % r).Interior.Color = GREEN
        ws.Range("E%d" % r).Formula = "=D%d*$D$17" % r
        ws.Range("F%d" % r).Formula = "=$D$21"
        ws.Range("G%d" % r).Formula = "=E%d*F%d" % (r, r)
        ws.Range("H%d" % r).Value = cpat
        ws.Range("H%d" % r).Interior.Color = GREEN
        ws.Range("I%d" % r).Formula = "=G%d-H%d" % (r, r)
        ws.Range("J%d" % r).Value = "egy.mit.ef.%s.ind" % f if f not in ("bio", "ren") else ""
        ws.Range("K%d" % r).Value = note
    r1 = r0 + len(FUELS) - 1
    ws.Range("D%d:D%d" % (r0, r1)).NumberFormat = "0.000000"
    ws.Range("E%d:H%d" % (r0, r1)).NumberFormat = "0.0000"
    ws.Range("I%d:I%d" % (r0, r1)).NumberFormat = "0.0E+00"
    ws.Range("A16").Value = "Parameters"
    ws.Range("A16").Font.Bold = True
    params = (
        (17, "GJ per ktoe", GJ_PER_KTOE, "CPAT Mitigation row 1035 (unit conversions)"),
        (18, "EFs selected (MTInputs EmissionsFactCO2)", "=" + mt_param("EmissionsFactCO2"),
         "Only the IIASA set is stored here (CPAT default IIASA*); another setting needs re-sourcing (REVIEW)"),
        (19, "Adjust non-Annex I EFs to inventories (MTInputs EFsAdjustmentNonAnnexI)",
         "=" + mt_param("EFsAdjustmentNonAnnexI"), "CPAT Mitigation row 976; Egypt is Non-Annex I (row 975)"),
        (20, "Adjustment factor, Egypt", ADJ,
         "CPAT Mitigation row 977: ratio of inventory (UNFCCC/IMF) to EF-estimated energy CO2, stored as a value"),
        (21, "Adjustment factor applied", '=IF(LEFT(D19,3)="Yes",D20,1)', "1 if the switch is off"),
    )
    for r, label, v, note in params:
        ws.Range("A%d" % r).Value = label
        if isinstance(v, str) and v.startswith("="):
            ws.Range("D%d" % r).Formula = v
        else:
            ws.Range("D%d" % r).Value = v
            ws.Range("D%d" % r).Interior.Color = GREEN
        ws.Range("K%d" % r).Value = note
    ws.Range("D20:D21").NumberFormat = "0.0000000000"
    ws.Range("K18").Interior.Color = REVIEW
    ws.Range("A23").Value = ("Check: CPAT irn 2022 = 232.657 ktoe coa x 3892.11 + 1669.057 ktoe nga x 2171.80 (/1e6) = "
                             "4.5304 MtCO2 = CPAT egy.mit.co2.ind.irn.1 (Mitigation row 6572).")
    return r0, r1


def build_emissions(wb):
    mi = wb.Worksheets("Mitigation_Industry")
    ws = wb.Worksheets.Add(After=mi)
    ws.Name = "Emissions_Industry"
    ws.Cells.Font.Name = "Arial"
    mi.Rows("1:2").Copy(ws.Rows(1))
    wb.Application.CutCopyMode = False
    for c in range(1, 36):
        ws.Columns(c).ColumnWidth = mi.Columns(c).ColumnWidth
    ws.Range("B2").Value = ("Mitigation Module - Egypt: Industry kernel - energy CO2 by sector and fuel "
                            "(Task F, Stream 2): eco2 = Mitigation_Industry ener (ktoe) x Data_EF EF (tCO2/ktoe) / 1e6")
    rows = {}   # (scenario, sector) -> (header row, first fuel row, total row)
    r = 3
    for s in (1, 2):
        copy_formats(mi.Rows(3), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d" % r).Value = ("Energy CO2 - %s (scenario %d), MtCO2; energy from Mitigation_Industry rows %d-%d"
                                     % ("Baseline" if s == 1 else "Carbon tax (policy scenario)", s,
                                        HEADERS[s][0], HEADERS[s][-1] + TOT))
        r += 2
        for h, sec in zip(HEADERS[s], SECTORS):
            hr = r
            copy_formats(mi.Rows(h), ws.Rows(hr))
            wb.Application.CutCopyMode = False
            ws.Range("B%d:J%d" % (hr, hr)).Value = (mi.Cells(h, 2).Value, sec, "Fuel Type", "Unit", "Variable Code",
                                                    "Input Code (Data_EF key)", "Output Code", "Scenario", s)
            ws.Range("K%d:AI%d" % (hr, hr)).Value = mi.Range("K%d:AI%d" % (h, h)).Value
            for i, f in enumerate(FUELS):
                rr, src = hr + 1 + i, h + ENER0 + i
                copy_formats(mi.Rows(src), ws.Rows(rr))
                wb.Application.CutCopyMode = False
                ws.Range("B%d" % rr).Value = f
                ws.Range("C%d" % rr).Formula = "=$C$%d" % hr
                ws.Range("D%d:F%d" % (rr, rr)).Value = (mi.Cells(src, 4).Value, "MtCO2", "eco2")
                ws.Range("G%d" % rr).Formula = '=Settings!$B$3&".mit.efc."&$B%d' % rr
                ws.Range("H%d" % rr).Formula = ('=Settings!$B$3&".mit.eco2."&$C$%d&"."&$B%d&".e."&$J$%d'
                                                % (hr, rr, hr))
                ws.Range("G%d:H%d" % (rr, rr)).Interior.Color = BLUE
                ws.Range("I%d" % rr).Value = "Mitigation_Industry row %d x Data_EF G / 1e6" % src
                for c in DATA_COLS:
                    ws.Cells(rr, c).FormulaR1C1 = ("=Mitigation_Industry!R%dC*INDEX(Data_EF!C7,MATCH(RC7,Data_EF!C1,0))"
                                                   "/1000000" % src)
            tr = hr + 1 + len(FUELS)
            copy_formats(mi.Rows(h + TOT), ws.Rows(tr))
            wb.Application.CutCopyMode = False
            ws.Range("B%d" % tr).Value = "tot"
            ws.Range("C%d" % tr).Formula = "=$C$%d" % hr
            ws.Range("D%d:F%d" % (tr, tr)).Value = ("Total", "MtCO2", "eco2")
            ws.Range("H%d" % tr).Formula = '=Settings!$B$3&".mit.eco2."&$C$%d&".tot.e."&$J$%d' % (hr, hr)
            ws.Range("H%d" % tr).Interior.Color = BLUE
            for c in DATA_COLS:
                ws.Cells(tr, c).FormulaR1C1 = "=SUM(R[-%d]C:R[-1]C)" % len(FUELS)
            ws.Range("%s%d:%s%d" % (col(DATA_COLS[0]), hr + 1, col(DATA_COLS[-1]), tr)).NumberFormat = "0.0000"
            rows[(s, sec)] = (hr, hr + 1, tr)
            r = tr + 2
        r += 1
    ws.Activate()
    wb.Application.ActiveWindow.FreezePanes = False
    ws.Range("L6").Select()
    wb.Application.ActiveWindow.FreezePanes = True
    return rows


def update_scenarios(wb, rows):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task F (stream 2, %s): energy CO2 codes egy.mit.eco2.<sector>.<fuel>.e.<s> and .tot (sheet "
        "Emissions_Industry, MtCO2, rows %d-%d scenario 1 and %d-%d scenario 2) = ener x EF; EF codes "
        "egy.mit.efc.<fuel> (sheet Data_EF, tCO2/ktoe = IIASA tCO2/GJ x 41868 x CPAT inventory adjustment 0.92962; "
        "bio and ren = 0)." % (VER, rows[(1, "irn")][0], rows[(1, "mch")][2], rows[(2, "irn")][0], rows[(2, "mch")][2]))


def update_check(wb, rows, ef_rows):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task F (Stream2 v1): energy CO2 by sector and fuel (expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff|",
                                              "rel. to max |CPAT|", "", "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
    maxabs = lambda r: "=MAX(MAX(%s%d:%s%d),-MIN(%s%d:%s%d))" % (L0, r, L1, r, L0, r, L1, r)
    r = hdr + 1
    # (a) totals = sum of fuel rows
    for s in (1, 2):
        for sec in SECTORS:
            h, f0, t = rows[(s, sec)]
            ws.Range("A%d" % r).Value = "eco2 total - sum of fuel rows (%s, Emissions_Industry row %d)" % (sec, t)
            ws.Range("D%d:E%d" % (r, r)).Value = (s, "diff")
            for c in DATA_COLS:
                L = col(c)
                ws.Cells(r, c).Formula = "=%s%s%d-SUM(%s%s%d:%s%d)" % (EI, L, t, EI, L, f0, L, f0 + len(FUELS) - 1)
            ws.Range("F%d" % r).Formula = maxabs(r)
            ws.Range("I%d" % r).Value = 0
            r += 1
    # (b) EF identity
    ws.Range("A%d" % r).Value = "Data_EF EF used - CPAT EF used (tCO2/ktoe; Mitigation rows 6516-6552 col B), max over fuels"
    ws.Range("E%d" % r).Value = "diff"
    ws.Range("F%d" % r).Formula = "=MAX(MAX(Data_EF!I%d:I%d),-MIN(Data_EF!I%d:I%d))" % (ef_rows * 2)
    ws.Range("I%d" % r).Value = 0
    r += 1
    # (c) baseline sector CO2 vs CPAT
    for sec in SECTORS:
        h, f0, t = rows[(1, sec)]
        code_c, code_k = "egy.mit.co2.ind.%s.1" % sec, "egy.mit.eco2.%s.tot.e.1" % sec
        ws.Range("A%d:E%d" % (r, r)).Value = (code_c, "tot", sec, 1, "CPAT")
        ws.Range("%s%d:%s%d" % (L0, r, L1, r)).Value = CPAT_CO2[sec]
        ws.Range("J%d" % r).Value = "CPAT_1_0pre_456 Mitigation row %d, mt.CO2" % CPAT_ROW[sec]
        ws.Range("A%d:E%d" % (r + 1, r + 1)).Value = (code_k, "tot", sec, 1, "kernel")
        for c in DATA_COLS:
            L = col(c)
            ws.Cells(r + 1, c).Formula = "=INDEX(%s%s:%s,MATCH($A%d,%s$H:$H,0))" % (EI, L, L, r + 1, EI)
            ws.Cells(r + 2, c).Formula = "=%s%d-%s%d" % (L, r + 1, L, r)
        ws.Range("A%d:E%d" % (r + 2, r + 2)).Value = (code_k, "tot", sec, 1, "diff")
        ws.Range("F%d" % (r + 2)).Formula = maxabs(r + 2)
        ws.Range("G%d" % (r + 2)).Formula = "=IF(MAX(%s%d:%s%d)=0,0,F%d/MAX(%s%d:%s%d))" % (L0, r, L1, r, r + 2, L0, r, L1, r)
        ws.Range("G%d" % (r + 2)).NumberFormat = "0.0E+00"
        ws.Range("I%d" % (r + 2)).Value = 0
        r += 3
    # (d) policy CO2 <= baseline CO2 when the explicit carbon price is positive
    for sec in SECTORS:
        _, a0, _ = rows[(1, sec)]
        _, b0, _ = rows[(2, sec)]
        a1, b1 = a0 + len(FUELS) - 1, b0 + len(FUELS) - 1
        ws.Range("A%d" % r).Value = ("Scenario 2 eco2 > scenario 1 eco2 while explicit carbon price (block o1, row %d) > 0 "
                                     "(count of fuel-years, %s)" % (PRICE_ROW, sec))
        ws.Range("D%d:E%d" % (r, r)).Value = (2, "count")
        for c in DATA_COLS:
            L = col(c)
            ws.Cells(r, c).Formula = "=SUMPRODUCT((%s%s%d:%s%d>%s%s%d:%s%d+1E-9)*(%s%s%d>0))" % (
                EI, L, b0, L, b1, EI, L, a0, L, a1, MI, L, PRICE_ROW)
        ws.Range("F%d" % r).Formula = "=SUM(%s%d:%s%d)" % (L0, r, L1, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task F)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    ws.Range("A%d" % (r + 1)).Value = ("Note: the CPAT comparison is exact only if the baseline ener rows match CPAT "
                                       "(Check!F2 = 0); energy CO2 here is combustion CO2 only (no process emissions, "
                                       "Task G).")
    return r0 + 1


def update_settings(wb, chk_row, rows):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: Task F section (row %d)." % (VER, chk_row)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task F (Stream 2, branched from v0.7): energy CO2 by sector x fuel x scenario. NEW SHEETS ONLY: Data_EF "
        "(after Data_Elast; IIASA industry fuel EFs from CPAT_1_0pre_456 Mitigation rows 1001-1009 x 41868 GJ/ktoe x "
        "inventory adjustment 0.92962 (row 977, applied while MTInputs EFsAdjustmentNonAnnexI = Yes); codes "
        "egy.mit.efc.<fuel>, tCO2/ktoe; bio and ren = 0 as in CPAT) and Emissions_Industry (after "
        "Mitigation_Industry; same columns/years; eco2 = ener x EF / 1e6, MtCO2; codes egy.mit.eco2.<sector>.<fuel>.e.<s>"
        " and .tot; scenario 1 rows %d-%d, scenario 2 rows %d-%d). Chosen because Mitigation_Industry has no spare "
        "rows after the sector sections (single blank separators) and the CBAM spare rows are used by Stream 1; "
        "references into Mitigation_Industry are by cell (ener rows h+49..h+58 for headers 5/65/125/185 and "
        "404/464/524/584). Merge = add the two sheets, this row, the Check section, the Scenarios note, MTInputs "
        "green on rows EmissionsFactCO2 / EFsAdjustmentNonAnnexI, and Data_EF / Emissions_Industry entries in the "
        "Sheets list above (not inserted here to keep row 25 fixed). No existing value changes. IPPU/process "
        "reconciliation is Task G."
        % (rows[(1, "irn")][0], rows[(1, "mch")][2], rows[(2, "irn")][0], rows[(2, "mch")][2]))
    ws.Range("C%d" % (last + 1)).WrapText = False


def colour_mtinputs(wb):
    ws = wb.Worksheets("MTInputs")
    want = {"EmissionsFactCO2", "EFsAdjustmentNonAnnexI"}
    for r in range(8, 416):
        if ws.Cells(r, 8).Value in want:
            ws.Range("B%d:H%d" % (r, r)).Interior.Color = GREEN
            want.discard(ws.Cells(r, 8).Value)
    if want:
        raise ValueError("MTInputs parameters not found: %s" % want)


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(DST, 0, False)
        xl.Calculation = -4135
        check_layout(wb)
        ef_rows = build_data_ef(wb)
        rows = build_emissions(wb)
        colour_mtinputs(wb)
        update_scenarios(wb, rows)
        chk = update_check(wb, rows, ef_rows)
        update_settings(wb, chk, rows)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        for b in list(xl.Workbooks):
            b.Close(False)
        xl.Quit()
        os.remove(DST)
        raise
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
    print("saved", DST)


if __name__ == "__main__":
    main()
