"""Build CPAT_Industry_Kernel_Egypt_v0.14.xlsx from v0.13 (Task M: links to main CPAT and Table 2 assembly).

Gap item 5.6 (EgyptTaskReference.md): national GHG for coverage %, total revenue, deaths avoided and recycling
effects, then Table 2. Method follows the TASK-2b ad-hoc rebuild (egypt/supporting/AdHocRebuild, MethodologyNote
v0.1) but uses the kernel's own responses instead of the rebuild's kappa / D_obr / F_fund corrections.

New sheet CPAT_National: CPAT national outputs of runs EG1-EG4 (cpat_outputs_egypt_2022_2041.csv), stored values.
New sheet Table2_Industry (Mitigation layout, data L..AE = 2022-2041):
  A  bundle -> CPAT run mapping (1A, 2A -> EG1; 2B -> EG2; 3A-3C -> EG3; LEGACY -> '-', not a Table 2 bundle) and the
     active bundle (Scenarios row 27: scope column L, process flag G, revenue use K); summary year selector.
  B  CPAT national series of the mapped run (0 for '-') and deltas policy - baseline.
  C  kernel series: energy CO2 of the kernel sectors (IPPU_Industry), kernel IPPU (Task G), non-kernel industry,
     block process emissions, revenue, rebates (Task I) and fund (Task J).
  D  Table 2 columns per year: J coverage, K reduction, L % of GHG, M/N/O/T/U block metrics, P net revenue, Q/R deaths
     avoided, AR net new revenue change. The composition is in deltas (the csv run and the kernel's stored CPAT
     baseline are different data vintages; levels are not mixed):
       K  = dGHG_CPAT + (dIPPU_kernel - dIPPU_CPAT) + (dInd_kernel - dInd_CPAT)
       dInd_kernel = d energy CO2 of kernel sectors + non-kernel industry change (All sectors: non-kernel baseline x
                     CPAT run's industry ratio; Industry only: 0, unpriced)
  E  summary at the selected year: live row for the active bundle + stored snapshot of all bundles (2030, written
     by this build) + rebuild v0.1 K for comparison.
Check: section 'Task M'; Settings version-log row; Scenarios note.
No existing formula changes: all prior sheets identical to v0.13 for every bundle (asserted).

Run with Excel installed (from this folder):  python build_v0_14.py
"""
import csv
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLUE, DATA_COLS, GREEN, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.13.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.13.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.14.xlsx")
CSV = os.path.normpath(os.path.join(HERE, "..", "..", "..", "egypt", "supporting", "AdHocRebuild",
                                    "cpat_outputs_egypt_2022_2041.csv"))
VER = "v0.14"
MI, IP, CN, T2 = "Mitigation_Industry!", "IPPU_Industry!", "CPAT_National!", "Table2_Industry!"
RB, FD, SC = "Rebate_Industry!", "Fund_Industry!", "Scenarios!"
BUNDLES = ("LEGACY", "1A", "2A", "2B", "3A", "3B", "3C")
RUN = {"LEGACY": "-", "1A": "EG1", "2A": "EG1", "2B": "EG2", "3A": "EG3", "3B": "EG3", "3C": "EG3"}
REBUILD_K = {"1A": -31.6, "2A": -27.7, "2B": -29.3, "3A": -18.0, "3B": -12.8, "3C": -22.4}
SCENS = ("EG1", "EG2", "EG3", "EG4")
BLOCK = {1: 245, 2: 644}
O_REVP, O_EMIS, O_EMISP = 138, 135, 145
O_MET = {"cbcov": 39, "cbintch": 41, "cbqch": 62, "emrt": 83, "cbobchu": 127}
IP_ROWS = {"ind": 11, "eco2.1": 17, "eco2.2": 34, "kern.1": 26, "kern.2": 43}
SC_BUNDLE0, SC_ACTIVE, SC_CPTRAJ, SC_ENCOV = 17, 27, 42, 45
RB_OBR, FD_FUND = 217, 329
L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
C30 = DATA_COLS[8]


def o(s, n):
    return BLOCK[s] + n


def unit_of(code):
    if ".rev." in code:
        return "$bn"
    if "cptraj" in code:
        return "$/tCO2"
    if code.startswith("egy.air.ada") or code == "egy.air.mort":
        return "deaths"
    if code.startswith("egy.air"):
        return "csv"
    return "MtCO2e" if ".ghg." in code else "MtCO2"


def read_csv():
    with open(CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    years = [str(y) for y in range(2022, 2042)]
    out = sorted(rows, key=lambda x: (x["code"], x["scenario"]))
    if len(out) != 128 or any(x["key"] != x["code"] + "|" + x["scenario"] for x in out):
        raise ValueError("unexpected csv layout")
    return [(x["key"], x["code"], x["scenario"], x["label"], [float(x[y]) for y in years]) for x in out]


# ----------------------------------------------------------------------------------------------- layout guards
def check_layout(wb):
    names = [ws.Name for ws in wb.Worksheets]
    if "CPAT_National" in names or "Table2_Industry" in names:
        raise ValueError("Task M sheets already exist")
    st = wb.Worksheets("Settings")
    if st.Range("B10").Value != "LEGACY" or st.Range("A25").Value != "Check" or st.Range("B3").Value != "egy":
        raise ValueError("Settings layout / bundle changed")
    sc = wb.Worksheets("Scenarios")
    for i, bd in enumerate(BUNDLES):
        if sc.Cells(SC_BUNDLE0 + i, 2).Value != bd:
            raise ValueError("Scenarios bundle row moved: %s" % bd)
    if "Energy scope" not in str(sc.Range("L16").Value) or sc.Range("G16").Value != "Process emissions covered (0/1)":
        raise ValueError("Scenarios bundle columns moved")
    if sc.Range("B%d" % SC_CPTRAJ).Value != "cptraj" or sc.Range("B%d" % SC_ENCOV).Value != "encov" or \
            sc.Cells(41, C30).Value != 2030:
        raise ValueError("Scenarios series rows moved")
    mi = wb.Worksheets("Mitigation_Industry")
    for s in (1, 2):
        for code, n in list(O_MET.items()) + [("revp", O_REVP), ("emis", O_EMIS), ("emisp", O_EMISP)]:
            if mi.Cells(o(s, n), 8).Value != "egy.mit.%s.cbam.tot.%d" % (code, s):
                raise ValueError("block row o%d is %s" % (n, mi.Cells(o(s, n), 8).Value))
    ip = wb.Worksheets("IPPU_Industry")
    exp = {"ind": "egy.mit.co2.ind.1", "eco2.1": "egy.mit.eco2.kern.tot.e.1", "eco2.2": "egy.mit.eco2.kern.tot.e.2",
           "kern.1": "egy.mit.ghg.iprk.tot.1", "kern.2": "egy.mit.ghg.iprk.tot.2"}
    for k, r in IP_ROWS.items():
        if ip.Cells(r, 8).Value != exp[k]:
            raise ValueError("IPPU_Industry row %d moved" % r)
    if wb.Worksheets("Rebate_Industry").Cells(RB_OBR, 8).Value != "egy.mit.obr.cbam.tot.2" or \
            wb.Worksheets("Fund_Industry").Cells(FD_FUND, 8).Value != "egy.mit.fund.cbam.tot.2":
        raise ValueError("rebate / fund total rows moved")


# ----------------------------------------------------------------------------------------------- sheet helpers
def new_sheet(wb, name, after, title):
    mi = wb.Worksheets("Mitigation_Industry")
    ws = wb.Worksheets.Add(After=wb.Worksheets(after))
    ws.Name = name
    ws.Cells.Font.Name = "Arial"
    mi.Rows("1:2").Copy(ws.Rows(1))
    wb.Application.CutCopyMode = False
    for c in range(1, 36):
        ws.Columns(c).ColumnWidth = mi.Columns(c).ColumnWidth
    ws.Range("B2").Value = title
    years = mi.Range("K5:AI5").Value

    def band(r, text):
        copy_formats(mi.Rows(3), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d" % r).Value = text

    def header(r, s="", labels=("Item", "Sector", "Description", "Unit", "Source", "Input Code", "Output Code",
                                "Note")):
        copy_formats(mi.Rows(5), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d:J%d" % (r, r)).Value = tuple(labels) + (s,)
        ws.Range("K%d:AI%d" % (r, r)).Value = years

    def line(r, item, sec, desc, unit, src, code, note, fml=None, vals=None, fmt="0.0000", incode=None):
        ws.Range("B%d:F%d" % (r, r)).Value = (item, sec, desc, unit, src)
        if incode:
            ws.Range("G%d" % r).Value = incode
        if code:
            ws.Range("H%d" % r).Formula = '=Settings!$B$3&"%s"' % code
            ws.Range("H%d" % r).Interior.Color = BLUE
        ws.Range("I%d" % r).Value = note
        if vals is not None:
            ws.Range("%s%d:%s%d" % (L0, r, L1, r)).Value = vals
            ws.Range("%s%d:%s%d" % (L0, r, L1, r)).Interior.Color = GREEN
        else:
            for c in DATA_COLS:
                ws.Cells(r, c).Formula = fml(c)
        ws.Range("%s%d:%s%d" % (L0, r, L1, r)).NumberFormat = fmt

    return ws, band, header, line


def freeze(wb, ws, cell="L6"):
    ws.Activate()
    wb.Application.ActiveWindow.FreezePanes = False
    ws.Range(cell).Select()
    wb.Application.ActiveWindow.FreezePanes = True


# ----------------------------------------------------------------------------------------------- CPAT_National
def build_national(wb, data):
    ws, band, header, line = new_sheet(
        wb, "CPAT_National", "IPPU_Industry",
        "Main CPAT - Egypt: national outputs of the CPAT runs EG1-EG4 used by Task M (stored values)")
    band(3, "A. CPAT national outputs by run (cpat_outputs_egypt_2022_2041.csv, egypt/supporting/AdHocRebuild; "
            "EG1 = 1A/2A, EG2 = 2B, EG3 = 3A-3C, EG4 unused)")
    header(5, "", ("Key (code|run)", "Run", "Description", "Unit", "Source", "", "CPAT code", "Note"))
    r = 6
    for key, code, scen, label, vals in data:
        ws.Range("B%d:F%d" % (r, r)).Value = (key, scen, label, unit_of(code), "CPAT run %s (csv)" % scen)
        ws.Range("H%d" % r).Value = code
        ws.Range("H%d" % r).Interior.Color = BLUE
        ws.Range("%s%d:%s%d" % (L0, r, L1, r)).Value = vals
        ws.Range("%s%d:%s%d" % (L0, r, L1, r)).Interior.Color = GREEN
        ws.Range("%s%d:%s%d" % (L0, r, L1, r)).NumberFormat = "#,##0.000"
        r += 1
    ws.Range("B%d" % (r + 1)).Value = (
        "Note: different data vintage from the CPAT values stored in the kernel (IPPU_Industry section A, CPAT_1_0pre_456 "
        "xlsb): 2030 baseline IPPU 85.97 vs 87.67, industrial energy CO2 88.97 vs 75.38 MtCO2. Task M uses deltas "
        "(policy - baseline) of a run only; levels of the two sources are never mixed.")
    freeze(wb, ws)
    return {"r0": 6, "r1": r - 1, "keys": {d[0]: 6 + i for i, d in enumerate(data)}}


# ----------------------------------------------------------------------------------------------- Table2_Industry
NAT = (
    ("ghg1", "egy.mit.ghg.tot.inc.1", "Total GHG incl. LULUCF, baseline", "MtCO2e"),
    ("ghg2", "egy.mit.ghg.tot.inc.2", "Total GHG incl. LULUCF, policy", "MtCO2e"),
    ("ipr1", "egy.mit.ghg.ipr.tot.1", "IPPU GHG, baseline", "MtCO2e"),
    ("ipr2", "egy.mit.ghg.ipr.tot.2", "IPPU GHG, policy", "MtCO2e"),
    ("enr1", "egy.mit.co2.enr.tot.1", "Energy CO2, baseline", "MtCO2"),
    ("enr2", "egy.mit.co2.enr.tot.2", "Energy CO2, policy", "MtCO2"),
    ("ind1", "egy.mit.co2.ind.1", "Industry energy CO2, baseline", "MtCO2"),
    ("ind2", "egy.mit.co2.ind.2", "Industry energy CO2, policy", "MtCO2"),
    ("cptraj", "egy.mit.cptraj.2", "Carbon price trajectory of the run", "$/tCO2"),
    ("effcp", "egy.mit.eff.cptraj.2", "Effective (coverage-weighted) carbon price of the run", "$/tCO2"),
    ("rev1", "egy.mit.rev.new.usd.1", "Net new revenue, baseline", "$bn"),
    ("rev2", "egy.mit.rev.new.usd.2", "Net new revenue, policy", "$bn"),
    ("rcoa", "egy.mit.rev.new.coa.usd.2", "Carbon-tax revenue, coal", "$bn"),
    ("rdie", "egy.mit.rev.new.die.usd.2", "Carbon-tax revenue, diesel", "$bn"),
    ("rgso", "egy.mit.rev.new.gso.usd.2", "Carbon-tax revenue, gasoline", "$bn"),
    ("rlpk", "egy.mit.rev.new.lpk.usd.2", "Carbon-tax revenue, LPG/kerosene", "$bn"),
    ("rnga", "egy.mit.rev.new.nga.usd.2", "Carbon-tax revenue, natural gas", "$bn"),
    ("roil", "egy.mit.rev.new.oil.usd.2", "Carbon-tax revenue, oil products", "$bn"),
    ("d2464", "egy.air.ada.2464", "Air-pollution deaths avoided, 25-64", "deaths"),
    ("d65", "egy.air.ada.65", "Air-pollution deaths avoided, 65+", "deaths"),
    ("du24", "egy.air.ada.u24", "Air-pollution deaths avoided, under 24", "deaths"),
    ("mort", "egy.air.mort", "Air-pollution deaths, baseline total", "deaths"),
)
SUMMARY = (  # (Table 2 column, key, label, format)
    ("J", "J", "Coverage, % of GHG", "0.0"), ("K", "K", "Reduction, MtCO2e", "0.00"),
    ("L", "L", "Reduction, % of GHG", "0.00"), ("M", "M", "CBAM coverage, %", "0.0"),
    ("N", "N", "CBAM intensity change, %", "0.0"), ("O", "O", "CBAM obligation per t exported, %", "0.0"),
    ("P", "P", "Net revenue, $bn", "0.00"), ("Q", "Q", "Deaths avoided", "#,##0"),
    ("R", "R", "Deaths avoided, % of baseline", "0.00"), ("T", "T", "Block emissions change, %", "0.0"),
    ("U", "U", "Block output change, %", "0.0"), ("AR", "AR", "Net new revenue change, $bn", "0.00"),
)


def build_table2(wb, nat):
    ws, band, header, line = new_sheet(
        wb, "Table2_Industry", "CPAT_National",
        "Mitigation Module - Egypt: Industry kernel - links to main CPAT and Table 2 (Task M)")
    k = {}
    # A. mapping and active bundle
    band(3, "A. Bundle -> CPAT run mapping and active bundle")
    copy_formats(wb.Worksheets("Mitigation_Industry").Rows(5), ws.Rows(5))
    wb.Application.CutCopyMode = False
    ws.Range("B5:F5").Value = ("Bundle", "Name", "CPAT run", "Energy scope (Scenarios L)", "Note")
    for i, bd in enumerate(BUNDLES):
        r = 6 + i
        ws.Range("B%d" % r).Formula = "=%s$B$%d" % (SC, SC_BUNDLE0 + i)
        ws.Range("C%d" % r).Formula = "=%s$C$%d" % (SC, SC_BUNDLE0 + i)
        ws.Range("D%d" % r).Value = RUN[bd]
        ws.Range("D%d" % r).Interior.Color = TAN
        ws.Range("E%d" % r).Formula = "=%s$L$%d" % (SC, SC_BUNDLE0 + i)
        ws.Range("F%d" % r).Value = ("'-': regression bundle, not in Table 2 (national rows 0)" if bd == "LEGACY" else
                                     "TASK-2b run mapping (MethodologyNote v0.1)")
    m0, m1 = 6, 6 + len(BUNDLES) - 1
    r = m1 + 2
    act = (("bundle", "Active bundle", "=%s$B$%d" % (SC, SC_ACTIVE)),
           ("run", "CPAT run", "=INDEX($D$%d:$D$%d,MATCH(D%d,$B$%d:$B$%d,0))" % (m0, m1, r, m0, m1)),
           ("scope", "Energy scope", "=%s$L$%d" % (SC, SC_ACTIVE)),
           ("flag", "Process emissions covered (0/1)", "=%s$G$%d" % (SC, SC_ACTIVE)),
           ("use", "Revenue use (recycling)", "=%s$K$%d" % (SC, SC_ACTIVE)),
           ("year", "Summary year (section E)", 2030))
    for i, (key, label, f) in enumerate(act):
        ws.Range("C%d" % (r + i)).Value = label
        ws.Range("D%d" % (r + i)).Formula = f
        k[key] = "$D$%d" % (r + i)
    ws.Range("D%d" % (r + 5)).Interior.Color = TAN
    ws.Range("E%d" % (r + 5)).Value = "the stored snapshot (E) is for 2030; Check compares live = snapshot in 2030"
    run, scope, flag = k["run"], k["scope"], k["flag"]
    r += 8

    # B. CPAT national series of the mapped run
    band(r, "B. CPAT national series of the mapped run (CPAT_National; 0 when the run is '-')")
    header(r + 2)
    r += 3
    for key, code, desc, unit in NAT:
        line(r, "nat", "cpat", desc, unit, "CPAT_National", ".mit.t2.%s" % key, "", incode=code,
             fml=lambda c, rr=r: '=IF(%s="-",0,INDEX(%s%s$%d:%s$%d,MATCH($G%d&"|"&%s,%s$B$%d:$B$%d,0)))'
             % (run, CN, col(c), nat["r0"], col(c), nat["r1"], rr, run, CN, nat["r0"], nat["r1"]), fmt="#,##0.000")
        k[key] = r
        r += 1
    R = lambda key, c: "%s%d" % (col(c), k[key])
    derived = (
        ("revctax", "Carbon-tax receipts of the run, six fuels", "$bn", "rebuild CV_revctax",
         lambda c: "=SUM(%s:%s)" % (R("rcoa", c), R("roil", c))),
        ("deaths", "Deaths avoided of the run, all ages", "deaths", "rebuild CV_deaths",
         lambda c: "=SUM(%s:%s)" % (R("d2464", c), R("du24", c))),
        ("dghg", "dGHG of the run (policy - baseline)", "MtCO2e", "", lambda c: "=%s-%s" % (R("ghg2", c),
                                                                                         R("ghg1", c))),
        ("dipr", "dIPPU of the run", "MtCO2e", "CPAT proportional IPPU response",
         lambda c: "=%s-%s" % (R("ipr2", c), R("ipr1", c))),
        ("denr", "dEnergy CO2 of the run", "MtCO2", "", lambda c: "=%s-%s" % (R("enr2", c), R("enr1", c))),
        ("dind", "dIndustry energy CO2 of the run", "MtCO2", "", lambda c: "=%s-%s" % (R("ind2", c), R("ind1", c))),
        ("drev", "dNet new revenue of the run", "$bn", "", lambda c: "=%s-%s" % (R("rev2", c), R("rev1", c))),
    )
    for key, desc, unit, note, f in derived:
        line(r, "der", "cpat", desc, unit, "", ".mit.t2.%s" % key, note, fml=f, fmt="#,##0.000")
        k[key] = r
        r += 1
    r += 2

    # C. kernel series
    band(r, "C. Kernel series (active bundle; scenario 1 = baseline, 2 = policy)")
    header(r + 2)
    r += 3
    kern = (
        ("ek1", "Energy CO2, kernel sectors, baseline", "MtCO2", "IPPU_Industry row %d" % IP_ROWS["eco2.1"],
         lambda c: "=%s%s%d" % (IP, col(c), IP_ROWS["eco2.1"])),
        ("ek2", "Energy CO2, kernel sectors, policy", "MtCO2", "IPPU_Industry row %d" % IP_ROWS["eco2.2"],
         lambda c: "=%s%s%d" % (IP, col(c), IP_ROWS["eco2.2"])),
        ("dek", "dEnergy CO2, kernel sectors", "MtCO2", "", lambda c: "=%s-%s" % (R("ek2", c), R("ek1", c))),
        ("indx", "Industry energy CO2, CPAT baseline (kernel's stored CPAT vintage)", "MtCO2",
         "IPPU_Industry row %d" % IP_ROWS["ind"], lambda c: "=%s%s%d" % (IP, col(c), IP_ROWS["ind"])),
        ("nk1", "Non-kernel industry energy CO2, baseline", "MtCO2", "", lambda c: "=%s-%s" % (R("indx", c),
                                                                                             R("ek1", c))),
        ("dnk", "dNon-kernel industry energy CO2", "MtCO2", "",
         lambda c: '=IF(%s="All sectors",IF(%s=0,0,%s*(%s/%s-1)),0)' % (scope, R("ind1", c), R("nk1", c),
                                                                       R("ind2", c), R("ind1", c))),
        ("dik", "dIndustry energy CO2, kernel composition", "MtCO2", "", lambda c: "=%s+%s" % (R("dek", c),
                                                                                             R("dnk", c))),
        ("ik1", "IPPU, kernel, baseline", "MtCO2e", "IPPU_Industry row %d" % IP_ROWS["kern.1"],
         lambda c: "=%s%s%d" % (IP, col(c), IP_ROWS["kern.1"])),
        ("ik2", "IPPU, kernel, policy", "MtCO2e", "IPPU_Industry row %d" % IP_ROWS["kern.2"],
         lambda c: "=%s%s%d" % (IP, col(c), IP_ROWS["kern.2"])),
        ("dik_p", "dIPPU, kernel", "MtCO2e", "", lambda c: "=%s-%s" % (R("ik2", c), R("ik1", c))),
        ("bp1", "Block process emissions (np + no), baseline", "MtCO2e", "Mitigation_Industry o145 (block 1)",
         lambda c: "=%s%s%d" % (MI, col(c), o(1, O_EMISP))),
        ("be1", "Block emissions (fuel + process), baseline", "MtCO2e", "Mitigation_Industry o135 (block 1)",
         lambda c: "=%s%s%d" % (MI, col(c), o(1, O_EMIS))),
        ("cp", "Carbon price on energy CO2, active bundle", "$/tCO2", "Scenarios row %d" % SC_CPTRAJ,
         lambda c: "=%s%s%d" % (SC, col(c), SC_CPTRAJ)),
        ("encov", "Share of industry energy CO2 covered", "share", "Scenarios row %d" % SC_ENCOV,
         lambda c: "=%s%s%d" % (SC, col(c), SC_ENCOV)),
        ("revp", "Block process revenue (np + no)", "$bn", "Mitigation_Industry o138 / 1000",
         lambda c: "=%s%s%d/1000" % (MI, col(c), o(2, O_REVP))),
        ("obr", "Output-based rebate cost", "$bn", "Rebate_Industry row %d / 1000" % RB_OBR,
         lambda c: "=%s%s%d/1000" % (RB, col(c), RB_OBR)),
        ("fund", "Abatement fund F", "$bn", "Fund_Industry row %d / 1000" % FD_FUND,
         lambda c: "=%s%s%d/1000" % (FD, col(c), FD_FUND)),
    )
    for key, desc, unit, src, f in kern:
        line(r, "kern", "ind", desc, unit, src, ".mit.t2.%s" % key, "", fml=f)
        k[key] = r
        r += 1
    r += 2

    # D. Table 2 columns per year
    band(r, "D. Table 2 columns per year, active bundle (composition in deltas; see Scenarios note and CAVEATS)")
    header(r + 2)
    r += 3
    nat0 = '%s="-"' % run
    tab = (
        ("ripr", "decomp", "K component: IPPU replacement (dIPPU kernel - dIPPU CPAT)", "MtCO2e", "",
         lambda c: "=IF(%s,0,%s-%s)" % (nat0, R("dik_p", c), R("dipr", c))),
        ("rind", "decomp", "K component: industry energy replacement (dInd kernel - dInd CPAT)", "MtCO2e", "",
         lambda c: "=IF(%s,0,%s-%s)" % (nat0, R("dik", c), R("dind", c))),
        ("K", "K", "[K] GHG reduction vs baseline = dGHG CPAT + IPPU + industry replacement", "MtCO2e", "",
         lambda c: "=IF(%s,0,%s+%s+%s)" % (nat0, R("dghg", c), R("ripr", c), R("rind", c))),
        ("J", "J", "[J] Coverage, % of baseline GHG (energy scope + flag x block process)", "%",
         "All sectors: CPAT energy CO2; Industry only: encov x kernel sectors",
         lambda c: '=IF(%s=0,0,100*(IF(%s="All sectors",%s,%s*%s)+%s*%s)/%s)'
         % (R("ghg1", c), scope, R("enr1", c), R("encov", c), R("ek1", c), flag, R("bp1", c), R("ghg1", c))),
        ("L", "L", "[L] Reduction, % of baseline GHG", "%", "",
         lambda c: "=IF(%s=0,0,100*%s/%s)" % (R("ghg1", c), R("K", c), R("ghg1", c))),
        ("M", "M", "[M] CBAM coverage", "%", "block cbcov o39", lambda c: "=100*%s%s%d" % (MI, col(c),
                                                                                      o(2, O_MET["cbcov"]))),
        ("N", "N", "[N] CBAM-sector intensity change", "%", "block cbintch o41",
         lambda c: "=100*%s%s%d" % (MI, col(c), o(2, O_MET["cbintch"]))),
        ("O", "O", "[O] CBAM obligation change per t exported", "%", "block cbobchu o127",
         lambda c: "=100*%s%s%d" % (MI, col(c), o(2, O_MET["cbobchu"]))),
        ("T", "T", "[T] Block emissions change (emrt / baseline block emissions)", "%", "block o83 / o135 (block 1)",
         lambda c: "=IF(%s=0,0,100*%s%s%d/%s)" % (R("be1", c), MI, col(c), o(2, O_MET["emrt"]), R("be1", c))),
        ("U", "U", "[U] Block output change", "%", "block cbqch o62",
         lambda c: "=100*%s%s%d" % (MI, col(c), o(2, O_MET["cbqch"]))),
        ("Pg", "P", "Gross carbon revenue (All: CPAT receipts adjusted to the kernel industry response; "
                    "Industry only: kernel sectors) + block process revenue", "$bn", "",
         lambda c: '=IF(%s,0,IF(%s="All sectors",%s+%s*%s*%s/1000,%s*%s*%s/1000)+%s)'
         % (nat0, scope, R("revctax", c), R("cp", c), R("encov", c), R("rind", c), R("cp", c), R("encov", c),
            R("ek2", c), R("revp", c))),
        ("P", "P", "[P] Net revenue = gross - output-based rebates - abatement fund", "$bn", "",
         lambda c: "=IF(%s,0,%s-%s-%s)" % (nat0, R("Pg", c), R("obr", c), R("fund", c))),
        ("denk", "Q", "dEnergy CO2, kernel composition (CPAT run with kernel industry response)", "MtCO2", "",
         lambda c: "=IF(%s,0,%s-%s+%s)" % (nat0, R("denr", c), R("dind", c), R("dik", c))),
        ("Q", "Q", "[Q] Air-pollution deaths avoided (CPAT deaths x dEnergy CO2 kernel / CPAT)", "deaths", "",
         lambda c: "=IF(%s=0,0,%s*%s/%s)" % (R("denr", c), R("deaths", c), R("denk", c), R("denr", c))),
        ("R", "R", "[R] Deaths avoided, % of baseline deaths", "%", "",
         lambda c: "=IF(%s=0,0,100*%s/%s)" % (R("mort", c), R("Q", c), R("mort", c))),
        ("AR", "AR", "[AR] Net new revenue change (CPAT dRev - CPAT carbon-tax receipts + P)", "$bn",
         "recycling: revenue use in section A", lambda c: "=IF(%s,0,%s-%s+%s)" % (nat0, R("drev", c),
                                                                                 R("revctax", c), R("P", c))),
    )
    for key, item, desc, unit, note, f in tab:
        line(r, item, "t2", desc, unit, "", ".mit.t2.col.%s" % key.lower(), note, fml=f)
        k[key] = r
        if key in ("K", "P"):
            ws.Range("B%d:AI%d" % (r, r)).Font.Bold = True
        r += 1
    r += 2

    # E. summary
    band(r, "E. Table 2 at the summary year: live row (active bundle) and stored snapshot of all bundles (2030)")
    r += 2
    copy_formats(wb.Worksheets("Mitigation_Industry").Rows(5), ws.Rows(r))
    wb.Application.CutCopyMode = False
    heads = ["Bundle", "CPAT run", "Energy scope", "Revenue use"] + ["[%s] %s" % (s[0], s[2]) for s in SUMMARY] + \
            ["Rebuild v0.1 K", "K - rebuild K"]
    for j, h in enumerate(heads):
        ws.Cells(r, 2 + j).Value = h
        ws.Cells(r, 2 + j).WrapText = True
    k["sum_hdr"] = r
    r += 1
    k["live"] = r
    ws.Range("A%d" % r).Value = "live"
    ws.Range("B%d" % r).Formula = "=" + k["bundle"]
    ws.Range("C%d" % r).Formula = "=" + run
    ws.Range("D%d" % r).Formula = "=" + scope
    ws.Range("E%d" % r).Formula = "=" + k["use"]
    yr = "MATCH(%s,$%s$%d:$%s$%d,0)" % (k["year"], L0, k["ripr"] - 1, L1, k["ripr"] - 1)  # section D year header
    for j, (_, key, _, fmt) in enumerate(SUMMARY):
        c = 6 + j
        ws.Cells(r, c).Formula = "=INDEX($%s$%d:$%s$%d,%s)" % (L0, k[key], L1, k[key], yr)
        ws.Cells(r, c).NumberFormat = fmt
    ws.Range("B%d:%s%d" % (r, col(6 + len(SUMMARY) - 1), r)).Font.Bold = True
    r += 1
    k["snap0"] = r
    for bd in BUNDLES[1:]:
        ws.Range("A%d" % r).Value = "stored"
        ws.Range("B%d" % r).Value = bd
        for j, (_, _, _, fmt) in enumerate(SUMMARY):
            ws.Cells(r, 6 + j).NumberFormat = fmt
        kc = 6 + len(SUMMARY)
        ws.Cells(r, kc).Value = REBUILD_K[bd]
        ws.Cells(r, kc).Interior.Color = GREEN
        ws.Cells(r, kc + 1).Formula = "=%s%d-%s%d" % (col(6 + 1), r, col(kc), r)
        ws.Cells(r, kc + 1).NumberFormat = "0.00"
        r += 1
    k["snap1"] = r - 1
    ws.Range("B%d" % (r + 1)).Value = (
        "Stored rows written by build_v0_14.py (2030, each bundle active in turn; Manual inputs at default). Rebuild v0.1 "
        "K: TASK-2b ad-hoc rebuild, MethodologyNote v0.1 section 3 (Mode REBUILD, Conv NOPHASE). Recycling effects: "
        "CPAT csv has no recycling outputs for the runs; reported as revenue use + AR only.")
    freeze(wb, ws)
    return k


def fill_snapshot(wb, k):
    xl, st, ws = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Table2_Industry")
    n = len(SUMMARY)
    for i, bd in enumerate(BUNDLES[1:]):
        st.Range("B10").Value = bd
        xl.CalculateFull()
        r = k["snap0"] + i
        ws.Range("C%d:E%d" % (r, r)).Value = ws.Range("C%d:E%d" % (k["live"], k["live"])).Value
        ws.Range("F%d:%s%d" % (r, col(5 + n), r)).Value = ws.Range("F%d:%s%d" % (k["live"], col(5 + n),
                                                                             k["live"])).Value
        ws.Range("C%d:%s%d" % (r, col(5 + n), r)).Interior.Color = GREEN
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()


# ----------------------------------------------------------------------------------------------- Check, notes
def update_check(wb, k, nat):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task M (%s): links to main CPAT and Table 2" % VER
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 3
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    maxabs = lambda rr: "=MAX(MAX(%s%d:%s%d),-MIN(%s%d:%s%d))" % (L0, rr, L1, rr, L0, rr, L1, rr)
    T = lambda key, c: "%s%s%d" % (T2, col(c), k[key])
    run = T2 + k["run"]
    base_codes = sorted({key.split("|")[0] for key in nat["keys"] if key.split("|")[0].endswith(".1")})

    def base_diff(c):
        terms = ["ABS(%s%s%d-%s%s%d)" % (CN, col(c), nat["keys"]["%s|%s" % (cd, s)], CN, col(c),
                                         nat["keys"]["%s|EG1" % cd]) for cd in base_codes for s in SCENS[1:]]
        return "=MAX(%s)" % ",".join(terms)

    n = len(SUMMARY)
    live = "%s$F$%d:$%s$%d" % (T2, k["live"], col(5 + n), k["live"])
    snap = "%s$F$%d:$%s$%d" % (T2, k["snap0"], col(5 + n), k["snap1"])
    snapb = "%s$B$%d:$B$%d" % (T2, k["snap0"], k["snap1"])
    items = [
        ("CPAT_National: baseline series (.1) identical across runs EG1-EG4", "-", "diff", base_diff, "max"),
        ("LEGACY (run '-'): national-linked Table 2 columns J, K, P, Q, AR = 0", 2, "diff",
         lambda c: '=IF(%s="-",ABS(%s)+ABS(%s)+ABS(%s)+ABS(%s)+ABS(%s),0)'
         % (run, T("J", c), T("K", c), T("P", c), T("Q", c), T("AR", c)), "max"),
        ("Industry only: industry replacement = dEnergy kernel sectors - dInd CPAT (non-kernel unpriced)", 2, "diff",
         lambda c: '=IF(%s%s="Industry only",%s-(%s-%s),0)' % (T2, k["scope"], T("rind", c), T("dek", c),
                                                               T("dind", c)), "max"),
        ("K - (dGHG CPAT + IPPU replacement + industry replacement)", 2, "diff",
         lambda c: "=%s-IF(%s=\"-\",0,%s+%s+%s)" % (T("K", c), run, T("dghg", c), T("ripr", c), T("rind", c)),
         "max"),
        ("Net revenue P > gross revenue (count of years)", 2, "count",
         lambda c: "=--(%s>%s+1E-9)" % (T("P", c), T("Pg", c)), "sum"),
    ]
    r = hdr + 1
    for label, s, kind, fml, agg in items:
        ws.Range("A%d" % r).Value = label
        ws.Range("D%d:E%d" % (r, r)).Value = (s, kind)
        for c in DATA_COLS:
            ws.Cells(r, c).Formula = fml(c)
        ws.Range("F%d" % r).Formula = maxabs(r) if agg == "max" else "=SUM(%s%d:%s%d)" % (L0, r, L1, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % r).Value = "Table 2 live row (summary year 2030) - stored snapshot row of the active bundle"
    ws.Range("D%d:E%d" % (r, r)).Value = (2, "diff")
    ws.Range("F%d" % r).Formula = (
        '=IF(OR(%s="-",%s%s<>2030),0,SUMPRODUCT(ABS(%s-INDEX(%s,MATCH(%s$B$%d,%s,0),0))))'
        % (run, T2, k["year"], live, snap, T2, k["live"], snapb))
    ws.Range("I%d" % r).Value = 0
    r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task M)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    r += 1
    ws.Range("A%d" % r).Value = "Information (not expected 0): 2030 K vs TASK-2b rebuild v0.1 (MtCO2e)"
    ws.Range("A%d" % r).Font.Bold = True
    r += 1
    for i, bd in enumerate(BUNDLES[1:]):
        sr = k["snap0"] + i
        ws.Range("A%d" % r).Value = "Info: %s K kernel composition - rebuild K" % bd
        ws.Range("D%d:E%d" % (r, r)).Value = (2, "value")
        ws.Range("F%d" % r).Formula = "=%s%s%d" % (T2, col(6 + len(SUMMARY) + 1), sr)
        ws.Range("F%d" % r).NumberFormat = "0.00"
        r += 1
    return r0 + 1


def update_scenarios(wb, k):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task M (%s): sheets CPAT_National (CPAT runs EG1-EG4, stored) and Table2_Industry. Bundle -> run mapping "
        "Table2_Industry D%d:D%d (1A/2A EG1, 2B EG2, 3A-3C EG3, LEGACY '-'); energy scope column L of this sheet. "
        "K = dGHG CPAT + (dIPPU kernel - dIPPU CPAT) + (dInd kernel - dInd CPAT), deltas only (csv vintage differs "
        "from the stored xlsb values). Table 2 live row %d, stored 2030 snapshot rows %d-%d."
        % (VER, 6, 6 + len(BUNDLES) - 1, k["live"], k["snap0"], k["snap1"]))


def update_settings(wb, chk, diff_note, stats):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: Task M section (row %d)." % (VER, chk)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task M: new sheets CPAT_National (CPAT national outputs of runs EG1-EG4 from the TASK-2b csv, stored) and "
        "Table2_Industry (bundle -> run mapping; CPAT series of the mapped run; kernel series; Table 2 columns J K L M "
        "N O P Q R T U AR per year; live summary row + stored 2030 snapshot of all bundles vs rebuild v0.1 K). "
        "Composition in deltas: K = dGHG CPAT + IPPU replacement (Task G kernel IPPU) + industry energy replacement "
        "(kernel sectors + non-kernel industry scaled with the run for All sectors). P = gross (CPAT receipts adjusted "
        "/ kernel sectors) + process revenue - rebates - fund. Q = CPAT deaths x dEnergy CO2 ratio. Check section row "
        "%d. 2030 K: %s. Regression vs v0.13: %s." % (chk, stats, diff_note))
    ws.Range("C%d" % (last + 1)).WrapText = False


# ----------------------------------------------------------------------------------------------- regression
SNAP = ("Mitigation_Industry", "Emissions_Industry", "Rebate_Industry", "Fund_Industry", "IPPU_Industry",
        "Data_Prices", "Scenarios", "Data_EF", "Data_Energy", "MTInputs")


def snapshot(wb):
    xl, st = wb.Application, wb.Worksheets("Settings")
    ck = wb.Worksheets("Check")
    last_ck = ck.Cells(ck.Rows.Count, 1).End(-4162).Row
    snap = {"last_ck": last_ck}
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        for n in SNAP:
            snap[(n, bd)] = wb.Worksheets(n).Range("A1:AI%d" % wb.Worksheets(n).UsedRange.Rows.Count).Value
        snap[("Check", bd)] = ck.Range("A1:AI%d" % last_ck).Value
        snap[("Manual inputs", bd)] = wb.Worksheets("Manual inputs").Range("A1:AE95").Value
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return snap


def n_empty(v):
    return v in (None, "")


def maxdiff(a, b):
    m, where = 0.0, None
    for i, (ra, rb) in enumerate(zip(a, b)):
        for j, (x, y) in enumerate(zip(ra, rb)):
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                if abs(x - y) > m:
                    m, where = abs(x - y), (i + 1, j + 1)
            elif x != y and not (n_empty(x) and n_empty(y)):
                if isinstance(x, str) and isinstance(y, str):
                    continue
                return float("inf"), (i + 1, j + 1, x, y)
    return m, where


def verify(wb, snap, k, chk):
    xl, st, ck = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Check")
    t2 = wb.Worksheets("Table2_Industry")
    max_rows = [r for r in range(1, ck.UsedRange.Rows.Count + 1) if str(ck.Cells(r, 1).Value).startswith("Max |")]
    legacy_only = {2, 3, 4, 584, 585, 1381}
    errs, stats = [], {}
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        bad = [r for r in max_rows if r not in legacy_only and
               (not isinstance(ck.Cells(r, 4).Value, (int, float)) or abs(ck.Cells(r, 4).Value) > 1e-9)]
        errs += ["%s: Check row %d %s = %s" % (bd, r, ck.Cells(r, 1).Value, ck.Cells(r, 4).Value) for r in bad]
        dmax = 0.0
        for n in SNAP + ("Check", "Manual inputs"):
            old = snap[(n, bd)]
            new = wb.Worksheets(n).Range("A1:%s%d" % ("AE" if n == "Manual inputs" else "AI", len(old))).Value
            d, w = maxdiff(old, new)
            if d > 1e-9:
                errs.append("%s: %s differs from v0.13 by %s at %s" % (bd, n, d, w))
            dmax = max(dmax, d)
        v = t2.Range("B%d:%s%d" % (k["live"], col(5 + len(SUMMARY)), k["live"])).Value[0]
        stats[bd] = dict(zip(["bundle", "run", "scope", "use"] + [s[0] for s in SUMMARY], v))
        s = stats[bd]
        print("bundle %-6s %s %-13s failing Max: %d | diff vs v0.13 %.3g | 2030 J %.1f K %.2f L %.2f M %.0f N %.1f "
              "O %.1f P %.2f Q %.0f R %.2f T %.1f U %.1f AR %.2f | Check M = %s"
              % (bd, s["run"], s["scope"], len(bad), dmax, s["J"], s["K"], s["L"], s["M"], s["N"], s["O"], s["P"],
                 s["Q"], s["R"], s["T"], s["U"], s["AR"], ck.Cells(chk, 4).Value))
        for ws in wb.Worksheets:
            vals = ws.UsedRange.Value
            if isinstance(vals, tuple):
                for i, rv in enumerate(vals):
                    for j, x in enumerate(rv):
                        if isinstance(x, int) and x < -2146820000:
                            errs.append("%s: error value %s!R%dC%d" % (bd, ws.Name, ws.UsedRange.Row + i,
                                                                       ws.UsedRange.Column + j))
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    if errs:
        raise ValueError("verification failed:\n" + "\n".join(errs[:40]))
    return stats


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    data = read_csv()
    shutil.copyfile(SRC, DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    wb = None
    try:
        wb = xl.Workbooks.Open(DST, 0, False)
        xl.Calculation = -4135
        check_layout(wb)
        print("snapshot v0.13 ...")
        snap = snapshot(wb)
        nat = build_national(wb, data)
        k = build_table2(wb, nat)
        xl.CalculateFull()
        fill_snapshot(wb, k)
        chk = update_check(wb, k, nat)
        update_scenarios(wb, k)
        xl.CalculateFull()
        print("verify ...")
        stats = verify(wb, snap, k, chk)
        txt = ", ".join("%s %.1f" % (bd, stats[bd]["K"]) for bd in BUNDLES[1:])
        update_settings(wb, chk, "all prior sheets identical for every bundle (max |diff| 0); new sheets only", txt)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        if os.environ.get("KEEP_DEBUG") and wb is not None:
            try:
                wb.SaveAs(os.path.join(HERE, "_debug_v0_14.xlsx"))
            except Exception:
                pass
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
