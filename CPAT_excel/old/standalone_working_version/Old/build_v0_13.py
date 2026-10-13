"""Build CPAT_Industry_Kernel_Egypt_v0.13.xlsx from v0.12 (Task G: IPPU double count and fuel-CO2 reconciliation).

Gap items 5.1 / 5.2 (EgyptTaskReference.md; MajorIssues #17, TASK-2a MI-1):
CPAT scales ALL IPPU emissions with industrial energy CO2 (CPAT_1_0pre_456 Mitigation rows 7553/7566-7569:
IPPU(t) = IPPU(t-1) x ind.CO2(t) / ind.CO2(t-1) from 2024; 2022-2023 inventory). The CBAM block models the
process emissions of the CBAM products explicitly (intensity + output response), so adding both double counts.

New sheet IPPU_Industry (after Fund_Industry, Mitigation layout, data L..AE = 2022-2041):
  A   CPAT baseline IPPU by gas (egy.mit.ghg.ipr.<gas>.1, rows 7566-7569, mt.CO2e) and industrial energy CO2
      (egy.mit.co2.ind.1, row 6590), stored as values (same source as the Task F check).
  B<s> per scenario s = 1, 2:
      eco2 kernel sectors      sum of Emissions_Industry sector totals (irn, cem, nfm, mch)
      ind CO2 implied          CPAT ind.CO2 baseline + (kernel eco2_s - kernel eco2_1)   (other sectors held)
      IPPU, CPAT method        CPAT proportional rule applied to 'ind CO2 implied' (scenario 1 = CPAT exactly)
      block np, no, process    CBAM block emisnp (o15), emisno (o16), emisp (o145)
      fp counted as IPPU       block post-policy fuel CO2 x fp share x flag (Manual inputs, ammonia feedstock)
      block IPPU               process + fp counted as IPPU
      non-block IPPU           s=1: CPAT IPPU - block IPPU; s=2: = s1 (IppuOther NONE) or s1 x CPAT ratio (CPAT)
      kernel IPPU              block IPPU + non-block IPPU   (replaces CPAT's proportional row)
      double-count correction  kernel IPPU - CPAT-method IPPU (0 in scenario 1)
  C<s> reconciliation (report, no calibration): block fuel CO2 net of fp counted as IPPU, by kernel sector
      (block mapping cpef: steel -> irn, clinker -> cem, ammonia/urea/AN -> mch, aluminium -> nfm), vs kernel sector
      energy CO2 (Task F); gap and ratio.
Manual inputs rows 84-95: switch IppuOther (E85) and per-product flag 'fp counted as IPPU' (E88:E95).
Check: section 'Task G' appended; Settings version-log row; Scenarios note.
No existing formula changes: all prior sheets identical to v0.12 for every bundle (asserted).

Run with Excel installed (from this folder):  python build_v0_13.py
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLUE, DATA_COLS, GREEN, REVIEW, TAN, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.12.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.12.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.13.xlsx")
VER = "v0.13"
MI, EI, IP, MN = "Mitigation_Industry!", "Emissions_Industry!", "IPPU_Industry!", "'Manual inputs'!"
BUNDLES = ("LEGACY", "1A", "2A", "2B", "3A", "3B", "3C")
BLOCK = {1: 245, 2: 644}
O_NP, O_NO, O_EF, O_FUEL, O_ERF, O_EMISP = 15, 16, 19, 85, 107, 145
O_CPEF = {"mch": 129, "irn": 130, "nfm": 131, "cem": 133}
SECTORS = ("irn", "cem", "nfm", "mch")
SEC_NAME = {"irn": "Iron and steel", "cem": "Cement", "nfm": "Non-ferrous metals", "mch": "Mining & chemicals"}
SEC_PROD = {"irn": (0, 2), "cem": (3, 3), "mch": (4, 6), "nfm": (7, 7)}     # contiguous product index ranges
EI_TOT = {1: {"irn": 15, "cem": 27, "nfm": 39, "mch": 51}, 2: {"irn": 66, "cem": 78, "nfm": 90, "mch": 102}}
PRODUCTS = ("stl.drg", "stl.scr", "stl.bof", "cmt.dry", "frt.amc", "frt.ure", "frt.nit", "alu.prp")
FP_FLAG = (0, 0, 0, 0, 1, 0, 0, 0)
FP_SRC = (
    "0: DRI reductant gas stays in energy CO2 (IEA/CPAT iron & steel final consumption; block steel fc+fp < CPAT irn)",
    "0: no fp", "0: BF-BOF coke/coal reductant counted with energy (not produced in Egypt, Q = 0)", "0: no fp",
    "1: SMR feedstock gas = IEA non-energy use (not in CPAT energy CO2; CPAT mch = 0) and IPCC 2B1 IPPU",
    "0: no fp", "0: no fp", "0: no fp")
L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
MAN_TITLE, MAN_SW, MAN_NOTE, MAN_HDR, MAN_R0 = 84, 85, 86, 87, 88

# CPAT_1_0pre_456 Mitigation, columns L..AE = 2022-2041 (scenario 1, mt.CO2e)
CPAT_IPR = (
    ("co2", 7566, "CO2", [39.4079807954547, 43.5850009187974, 50.91280409838354, 56.75741402705363, 55.1773166677362,
                          60.177140336361155, 62.59285088925789, 64.93258269722676, 67.14695788420454,
                          69.39700309387531, 71.66046847764495, 73.9668101898482, 76.3158077221917, 78.70721185504082,
                          81.1407441181349, 83.61609616553793, 86.13292906095532, 88.69087246971075,
                          91.28952375383629, 93.9284469668718]),
    ("ch4", 7567, "CH4", [0.1657182072, 0.165889424880574, 0.19377986951689197, 0.21602511350632642,
                          0.2100110849032105, 0.22904097719202224, 0.23823544377112332, 0.24714072668535048,
                          0.2555688881742523, 0.2641328137890065, 0.27274781809863574, 0.28152601454581,
                          0.29046656385111685, 0.29956851745654345, 0.30883081547289826, 0.3182522843008027,
                          0.3278316339105267, 0.33756745476656086, 0.34745821438342783, 0.35750225349977566]),
    ("n2o", 7568, "N2O", [0.9275, 0.9275, 1.083437531393625, 1.2078123299382222, 1.1741874528045189,
                          1.2805849830303275, 1.3319919232753465, 1.381782016338192, 1.428904488349799,
                          1.4767860275944065, 1.524953151586388, 1.5740326948460954, 1.6240199649004818,
                          1.6749096582917917, 1.7266958491014452, 1.7793719756488264, 1.8329308252826433,
                          1.8873645171859845, 1.9426644831196078, 1.9988214460310123]),
    ("fga", 7569, "F-gases", [11.6991507578219, 12.2250407871263, 14.280396777995145, 15.919735845486592,
                              15.476538546918793, 16.878925767037405, 17.55650198400422, 18.21276712523128,
                              18.833871321815767, 19.46498050803217, 20.099853883127096, 20.746753525351846,
                              21.4056175849227, 22.076376158888326, 22.758951139867214, 23.453256040729276,
                              24.159195794136256, 24.876666525901445, 25.60555530117394, 26.345739842492602]),
)
CPAT_IPR_TOT = [52.2003497604766, 56.90343113080428, 66.47041827728921, 74.10098731598477, 72.03805375236271,
                78.56569206362092, 81.71958024030857, 84.77427256548158, 87.66530258254436, 90.60290244329089,
                93.55802333045708, 96.56912242459195, 99.63591183586601, 102.75806618967749, 105.93522192257646,
                109.16697646621682, 112.45288731428474, 115.79247096756474, 119.18520175251327, 122.6305105088952]
CPAT_IND = [59.804205879849356, 48.931828957170424, 57.15857678914482, 63.720179343286866, 61.946242160955606,
            67.5594235630846, 70.27148351702611, 72.89824396717464, 75.38426956341291, 77.91034103351876,
            80.45147900929716, 83.04075320453173, 85.67791605988172, 88.36268778247626, 91.09475574047232,
            93.87377376132702, 96.69936132943994, 99.57110267900502, 102.4885457780901, 105.45120120012244]


def o(s, n):
    return BLOCK[s] + n


# ----------------------------------------------------------------------------------------------- layout guards
def check_layout(wb):
    names = [ws.Name for ws in wb.Worksheets]
    if "IPPU_Industry" in names:
        raise ValueError("IPPU_Industry already exists")
    st = wb.Worksheets("Settings")
    if st.Range("B10").Value != "LEGACY" or st.Range("A25").Value != "Check":
        raise ValueError("Settings layout / bundle changed")
    mi = wb.Worksheets("Mitigation_Industry")
    for s in (1, 2):
        if mi.Cells(BLOCK[s], 10).Value != s:
            raise ValueError("block %d header moved" % s)
        for n, code in ((O_NP, "emisnp"), (O_NO, "emisno"), (O_EMISP, "emisp")):
            v = str(mi.Cells(o(s, n), 8).Value)
            if ".mit.%s.cbam.tot" % code not in v:
                raise ValueError("block row o%d (scenario %d) is %s, expected %s" % (n, s, v, code))
        for sec, n in O_CPEF.items():
            if ".mit.cpef.ind.%s" % ("mac" if sec == "mac" else sec) not in str(mi.Cells(o(s, n), 8).Value):
                raise ValueError("cpef row for %s moved (scenario %d)" % (sec, s))
        hdr = [mi.Cells(o(s, O_EF) - 1, c).Value for c in (8, 9, 10, 11)]
        if hdr != ["S1.fc", "S1.fp", "S1.np", "S1.no"]:
            raise ValueError("EF header moved: %s" % hdr)
        if "Pre-policy Fuel" not in str(mi.Cells(o(s, O_FUEL) - 1, 3).Value) or \
                "Additional Emissions" not in str(mi.Cells(o(s, O_ERF) - 1, 3).Value):
            raise ValueError("fuel emission / fuel ER rows moved (scenario %d)" % s)
        for i in range(8):
            if mi.Cells(o(s, O_FUEL) + i, 6).Formula != "=H%d+I%d" % (o(s, O_EF) + i, o(s, O_EF) + i):
                raise ValueError("block F column is not fc + fp (row %d)" % (o(s, O_FUEL) + i))
    ei = wb.Worksheets("Emissions_Industry")
    for s, d in EI_TOT.items():
        for sec, r in d.items():
            if ei.Cells(r, 8).Value != "egy.mit.eco2.%s.tot.e.%d" % (sec, s):
                raise ValueError("Emissions_Industry total row moved: %d" % r)
    mn = wb.Worksheets("Manual inputs")
    for r in range(MAN_TITLE - 1, MAN_R0 + 9):
        if any(mn.Cells(r, c).Formula not in ("", None) for c in range(1, 32)):
            raise ValueError("Manual inputs row %d not empty" % r)


# ----------------------------------------------------------------------------------------------- Manual inputs
def build_manual(wb):
    ws = wb.Worksheets("Manual inputs")
    mi = wb.Worksheets("Mitigation_Industry")
    for src, dst in ((62, MAN_TITLE), (63, MAN_SW), (64, MAN_NOTE), (65, MAN_HDR)):
        copy_formats(ws.Rows(src), ws.Rows(dst))
    for i in range(8):
        copy_formats(ws.Rows(66), ws.Rows(MAN_R0 + i))
    wb.Application.CutCopyMode = False
    ws.Range("B%d" % MAN_TITLE).Value = "IPPU accounting (Task G, %s)" % VER
    ws.Range("D%d" % MAN_SW).Value = "Non-block IPPU response (IppuOther: NONE / CPAT)"
    ws.Range("E%d" % MAN_SW).Value = "NONE"
    ws.Range("E%d" % MAN_SW).Interior.Color = TAN
    ws.Range("E%d" % MAN_SW).Validation.Delete()
    ws.Range("E%d" % MAN_SW).Validation.Add(3, 1, 1, "NONE,CPAT")
    ws.Range("F%d" % MAN_SW).Value = (
        "NONE (default, MajorIssues #17): IPPU outside the CBAM block held at baseline in the policy scenario. CPAT: "
        "scaled with industrial energy CO2 as in CPAT. Code egy.in.ippuoth")
    ws.Range("D%d" % MAN_NOTE).Value = (
        "IPPU_Industry: kernel IPPU = block process (np + no, post-policy) + fp flagged below + non-block remainder "
        "(CPAT IPPU baseline - block IPPU baseline); replaces CPAT's proportional IPPU row")
    ws.Range("D%d:H%d" % (MAN_HDR, MAN_HDR)).Value = ("Product", "fp counted as IPPU (1/0)", "Source / rationale",
                                                     "Confidence", "Code")
    for i, p in enumerate(PRODUCTS):
        r = MAN_R0 + i
        ws.Range("D%d" % r).Value = mi.Cells(o(2, O_EF) + i, 4).Value
        ws.Range("E%d" % r).Value = FP_FLAG[i]
        ws.Range("E%d" % r).Interior.Color = TAN
        ws.Range("F%d" % r).Value = FP_SRC[i]
        ws.Range("G%d" % r).Value = "assumption"
        ws.Range("H%d" % r).Value = "egy.in.ippufp." + p
        ws.Range("H%d" % r).Interior.Color = BLUE
        for c in range(9, 32):
            ws.Cells(r, c).ClearContents()


# ----------------------------------------------------------------------------------------------- IPPU sheet
def build_ippu(wb):
    mi = wb.Worksheets("Mitigation_Industry")
    ws = wb.Worksheets.Add(After=wb.Worksheets("Fund_Industry"))
    ws.Name = "IPPU_Industry"
    ws.Cells.Font.Name = "Arial"
    mi.Rows("1:2").Copy(ws.Rows(1))
    wb.Application.CutCopyMode = False
    for c in range(1, 36):
        ws.Columns(c).ColumnWidth = mi.Columns(c).ColumnWidth
    ws.Range("B2").Value = ("Mitigation Module - Egypt: Industry kernel - IPPU accounting and fuel-CO2 reconciliation "
                            "(Task G): block process emissions replace CPAT's proportional IPPU row")
    years = mi.Range("K5:AI5").Value
    rows = {}

    def band(r, text):
        copy_formats(mi.Rows(3), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d" % r).Value = text

    def header(r, s):
        copy_formats(mi.Rows(5), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d:J%d" % (r, r)).Value = ("Item", "Sector", "Description", "Unit", "Source", "Input Code",
                                              "Output Code", "Note", s)
        ws.Range("K%d:AI%d" % (r, r)).Value = years

    def line(r, item, sec, desc, unit, src, code, note, fml=None, vals=None, fmt="0.0000"):
        ws.Range("B%d:F%d" % (r, r)).Value = (item, sec, desc, unit, src)
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

    # A. CPAT data
    band(3, "A. CPAT IPPU inventory and industrial energy CO2, baseline (CPAT_1_0pre_456 Mitigation, stored values)")
    header(5, "")
    r = 6
    for gas, row, name, vals in CPAT_IPR:
        line(r, gas, "ipr", "IPPU %s, CPAT baseline" % name, "MtCO2e", "CPAT Mitigation row %d" % row,
             ".mit.ghg.ipr.%s.1" % gas, "2022-23 inventory (rows 844-847); 2024+ x ind CO2 growth (row 7553)",
             vals=vals)
        rows["ipr." + gas] = r
        r += 1
    line(r, "tot", "ipr", "IPPU all gases, CPAT baseline", "MtCO2e", "CPAT Mitigation row 7570",
         ".mit.ghg.ipr.tot.1", "sum of the gas rows", fml=lambda c: "=SUM(%s%d:%s%d)" % (col(c), r - 4, col(c), r - 1))
    rows["ipr.tot"] = r
    r += 1
    line(r, "tot", "ind", "Industrial energy CO2 (all industry sectors), CPAT baseline", "MtCO2",
         "CPAT Mitigation row 6590", ".mit.co2.ind.1", "IPPU growth driver in CPAT", vals=CPAT_IND)
    rows["ind"] = r
    r += 3

    # B. IPPU accounting per scenario
    for s in (1, 2):
        b = BLOCK[s]
        band(r, "B%d. IPPU accounting - %s (scenario %d)" % (s, "Baseline" if s == 1 else "Policy scenario", s))
        header(r + 2, s)
        r += 3
        k = {}
        k["eco2"] = r
        line(r, "tot", "kern", "Energy CO2, kernel sectors (irn + cem + nfm + mch)", "MtCO2", "Emissions_Industry",
             ".mit.eco2.kern.tot.e.%d" % s, "Task F sector totals",
             fml=lambda c: "=" + "+".join("%s%s%d" % (EI, col(c), EI_TOT[s][x]) for x in SECTORS))
        r += 1
        k["ind"] = r
        line(r, "tot", "ind", "Industrial energy CO2 implied (CPAT baseline + kernel change)", "MtCO2",
             "row %d + kernel change" % rows["ind"], ".mit.co2.ind.k.%d" % s,
             "non-kernel industry sectors held at the CPAT baseline",
             fml=lambda c: "=%s%d+%s%d-%s%d" % (col(c), rows["ind"], col(c), r - 1, col(c), rows1["eco2"] if s == 2
                                                 else r - 1))
        r += 1
        k["cpat"] = r
        line(r, "tot", "ipr", "IPPU, CPAT method (all IPPU scaled with industrial energy CO2)", "MtCO2e",
             "CPAT Mitigation rows 7553, 7566-7570", ".mit.ghg.iprc.tot.%d" % s,
             "the row Task G replaces; scenario 1 = CPAT egy.mit.ghg.ipr.tot.1",
             fml=lambda c: ("=%s%d" % (col(c), rows["ipr.tot"]) if c < 14 else
                            "=IF(%s%d=0,0,%s%d*%s%d/%s%d)" % (col(c - 1), r - 1, col(c - 1), r, col(c), r - 1,
                                                             col(c - 1), r - 1)))
        r += 1
        for key, n, desc, code in (("np", O_NP, "Block process CO2 (np), post-policy", "emisnp"),
                                   ("no", O_NO, "Block non-CO2 process (no), post-policy", "emisno"),
                                   ("proc", O_EMISP, "Block process emissions (np + no), post-policy", "emisp")):
            k[key] = r
            line(r, "cbam", "ipr", desc, "MtCO2e", "Mitigation_Industry row %d (o%d)" % (o(s, n), n),
                 ".mit.ghg.ipr.blk%s.%d" % (key, s), "egy.mit.%s.cbam.tot.%d" % (code, s),
                 fml=lambda c, rr=o(s, n): "=%s%s%d" % (MI, col(c), rr))
            r += 1
        k["fp"] = r
        f0, f1, e0, e1, x0, x1 = (o(s, O_FUEL), o(s, O_FUEL) + 7, o(s, O_ERF), o(s, O_ERF) + 7, o(s, O_EF),
                                  o(s, O_EF) + 7)
        line(r, "cbam", "ipr", "Block fuel used in the process (fp) counted as IPPU", "MtCO2e",
             "Mitigation_Industry o85-o92 + o107-o114", ".mit.ghg.ipr.blkfp.%d" % s,
             "(fuel + fuel ER) x fp/(fc+fp) x flag (Manual inputs E%d:E%d)" % (MAN_R0, MAN_R0 + 7),
             fml=lambda c: ("=SUMPRODUCT((%s%s%d:%s%d+%s%s%d:%s%d)*%s$I$%d:$I$%d/(%s$F$%d:$F$%d+(%s$F$%d:$F$%d=0))"
                            "*%s$E$%d:$E$%d)" % (MI, col(c), f0, col(c), f1, MI, col(c), e0, col(c), e1, MI, x0, x1,
                                                 MI, f0, f1, MI, f0, f1, MN, MAN_R0, MAN_R0 + 7)))
        r += 1
        k["blk"] = r
        line(r, "cbam", "ipr", "Block IPPU (process + fp counted as IPPU)", "MtCO2e", "", ".mit.ghg.ipr.blk.%d" % s,
             "", fml=lambda c: "=%s%d+%s%d" % (col(c), k["proc"], col(c), k["fp"]))
        r += 1
        k["oth"] = r
        if s == 1:
            fml = lambda c: "=%s%d-%s%d" % (col(c), k["cpat"], col(c), k["blk"])
            note = "CPAT IPPU baseline - block IPPU baseline (about 40-60% of IPPU)"
        else:
            fml = lambda c: '=IF(%s$E$%d="CPAT",IF(%s%d=0,0,%s%d*%s%d/%s%d),%s%d)' % (
                MN, MAN_SW, col(c), rows1["cpat"], col(c), rows1["oth"], col(c), k["cpat"], col(c), rows1["cpat"],
                col(c), rows1["oth"])
            note = "IppuOther (Manual inputs E%d): NONE = baseline; CPAT = baseline x CPAT-method ratio" % MAN_SW
        line(r, "oth", "ipr", "Non-block IPPU (IPPU outside the CBAM block)", "MtCO2e", "", ".mit.ghg.ipr.oth.%d" % s,
             note, fml=fml)
        r += 1
        k["kern"] = r
        line(r, "tot", "ipr", "IPPU, kernel (block + non-block) - replaces the CPAT-method row", "MtCO2e", "",
             ".mit.ghg.iprk.tot.%d" % s, "use instead of egy.mit.ghg.ipr.tot.%d" % s,
             fml=lambda c: "=%s%d+%s%d" % (col(c), k["blk"], col(c), k["oth"]))
        ws.Range("B%d:AI%d" % (r, r)).Font.Bold = True
        r += 1
        k["dc"] = r
        line(r, "tot", "ipr", "Double-count correction: kernel IPPU - CPAT-method IPPU", "MtCO2e", "",
             ".mit.ghg.iprdc.tot.%d" % s, "add to CPAT total GHG to correct it (0 in scenario 1)",
             fml=lambda c: "=%s%d-%s%d" % (col(c), k["kern"], col(c), k["cpat"]))
        r += 1
        k["share"] = r
        line(r, "shr", "ipr", "Block IPPU share of CPAT-method IPPU", "share", "", ".mit.ghg.iprshb.tot.%d" % s, "",
             fml=lambda c: "=IF(%s%d=0,0,%s%d/%s%d)" % (col(c), k["cpat"], col(c), k["blk"], col(c), k["cpat"]),
             fmt="0.000")
        r += 3
        rows[("B", s)] = k
        if s == 1:
            rows1 = k

    # C. reconciliation
    for s in (1, 2):
        band(r, "C%d. Reconciliation: block product fuel CO2 vs kernel sector energy CO2 (Task F) - scenario %d"
             % (s, s))
        header(r + 2, s)
        r += 3
        k = {}
        for sec in SECTORS:
            i0, i1 = SEC_PROD[sec]
            f0, f1, e0, e1, x0, x1 = (o(s, O_FUEL) + i0, o(s, O_FUEL) + i1, o(s, O_ERF) + i0, o(s, O_ERF) + i1,
                                      o(s, O_EF) + i0, o(s, O_EF) + i1)
            line(r, "blk", sec, "%s: block fuel CO2 (fc + fp, post-policy) net of fp counted as IPPU" % SEC_NAME[sec],
                 "MtCO2", "Mitigation_Industry row %d (cpef)" % o(s, O_CPEF[sec]), ".mit.eco2.blk.%s.%d" % (sec, s),
                 "products: %s" % ", ".join(PRODUCTS[i0:i1 + 1]),
                 fml=lambda c: ("=%s%s%d-SUMPRODUCT((%s%s%d:%s%d+%s%s%d:%s%d)*%s$I$%d:$I$%d/(%s$F$%d:$F$%d+"
                                "(%s$F$%d:$F$%d=0))*%s$E$%d:$E$%d)"
                                % (MI, col(c), o(s, O_CPEF[sec]), MI, col(c), f0, col(c), f1, MI, col(c), e0, col(c),
                                   e1, MI, x0, x1, MI, f0, f1, MI, f0, f1, MN, MAN_R0 + i0, MAN_R0 + i1)))
            line(r + 1, "tot", sec, "%s: kernel sector energy CO2" % SEC_NAME[sec], "MtCO2",
                 "Emissions_Industry row %d" % EI_TOT[s][sec], "", "egy.mit.eco2.%s.tot.e.%d" % (sec, s),
                 fml=lambda c: "=%s%s%d" % (EI, col(c), EI_TOT[s][sec]))
            line(r + 2, "gap", sec, "%s: sector energy CO2 - block fuel CO2 (other products / non-CBAM)" % SEC_NAME[sec],
                 "MtCO2", "", ".mit.eco2.gap.%s.%d" % (sec, s), "negative = block exceeds the sector",
                 fml=lambda c: "=%s%d-%s%d" % (col(c), r + 1, col(c), r))
            line(r + 3, "rat", sec, "%s: block fuel CO2 / sector energy CO2" % SEC_NAME[sec], "ratio", "",
                 ".mit.eco2.rblk.%s.%d" % (sec, s), "0 if the sector has no energy CO2 (CPAT nfm, mch)",
                 fml=lambda c: "=IF(%s%d=0,0,%s%d/%s%d)" % (col(c), r + 1, col(c), r, col(c), r + 1), fmt="0.000")
            k[sec] = r
            r += 5
        rows[("C", s)] = k
        r += 1
    ws.Activate()
    wb.Application.ActiveWindow.FreezePanes = False
    ws.Range("L6").Select()
    wb.Application.ActiveWindow.FreezePanes = True
    return rows


# ----------------------------------------------------------------------------------------------- Check, notes
def update_check(wb, rows):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task G (%s): IPPU accounting and fuel-CO2 reconciliation" % VER
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 3
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    maxabs = lambda rr: "=MAX(MAX(%s%d:%s%d),-MIN(%s%d:%s%d))" % (L0, rr, L1, rr, L0, rr, L1, rr)
    b1, b2 = rows[("B", 1)], rows[("B", 2)]
    r = hdr + 1
    items = [
        ("IPPU CPAT method, scenario 1 - CPAT egy.mit.ghg.ipr.tot.1 (stored)", 1, "diff",
         lambda c: "=%s%s%d-%s%s%d" % (IP, col(c), b1["cpat"], IP, col(c), rows["ipr.tot"]), "max"),
        ("Kernel IPPU - CPAT-method IPPU, scenario 1", 1, "diff",
         lambda c: "=%s%s%d" % (IP, col(c), b1["dc"]), "max"),
    ]
    for s, b in ((1, b1), (2, b2)):
        items.append(("Block np + no - block process total (emisnp + emisno - emisp)", s, "diff",
                      lambda c, b=b: "=%s%s%d+%s%s%d-%s%s%d" % (IP, col(c), b["np"], IP, col(c), b["no"], IP, col(c),
                                                               b["proc"]), "max"))
        items.append(("Non-block IPPU < 0 (count of years)", s, "count",
                      lambda c, b=b: "=--(%s%s%d<-1E-9)" % (IP, col(c), b["oth"]), "sum"))
        items.append(("Kernel IPPU - (block + non-block)", s, "diff",
                      lambda c, b=b: "=%s%s%d-%s%s%d-%s%s%d" % (IP, col(c), b["kern"], IP, col(c), b["blk"], IP,
                                                               col(c), b["oth"]), "max"))
    items += [
        ("Block CO2 (np + fp as IPPU) > CPAT IPPU CO2, scenario 1 (count of years)", 1, "count",
         lambda c: "=--(%s%s%d+%s%s%d>%s%s%d+1E-9)" % (IP, col(c), b1["np"], IP, col(c), b1["fp"], IP, col(c),
                                                      rows["ipr.co2"]), "sum"),
        ("Block non-CO2 (no) > CPAT IPPU CH4 + N2O + F-gases, scenario 1 (count of years)", 1, "count",
         lambda c: "=--(%s%s%d>%s%s%d+%s%s%d+%s%s%d+1E-9)" % (IP, col(c), b1["no"], IP, col(c), rows["ipr.ch4"], IP,
                                                             col(c), rows["ipr.n2o"], IP, col(c), rows["ipr.fga"]),
         "sum"),
        ("Non-block IPPU scenario 2 - IppuOther rule (NONE: = s1; CPAT: s1 x CPAT ratio)", 2, "diff",
         lambda c: ('=%s%s%d-IF(%s$E$%d="CPAT",IF(%s%s%d=0,0,%s%s%d*%s%s%d/%s%s%d),%s%s%d)'
                    % (IP, col(c), b2["oth"], MN, MAN_SW, IP, col(c), b1["cpat"], IP, col(c), b1["oth"], IP, col(c),
                       b2["cpat"], IP, col(c), b1["cpat"], IP, col(c), b1["oth"])), "max"),
    ]
    for label, s, kind, fml, agg in items:
        ws.Range("A%d" % r).Value = label
        ws.Range("D%d:E%d" % (r, r)).Value = (s, kind)
        for c in DATA_COLS:
            ws.Cells(r, c).Formula = fml(c)
        ws.Range("F%d" % r).Formula = maxabs(r) if agg == "max" else "=SUM(%s%d:%s%d)" % (L0, r, L1, r)
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task G)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    # informational (not expected 0)
    r += 1
    ws.Range("A%d" % r).Value = "Information (not expected 0): reconciliation and double-count size"
    ws.Range("A%d" % r).Font.Bold = True
    r += 1
    for s in (1, 2):
        for sec in SECTORS:
            rc = rows[("C", s)][sec]
            ws.Range("A%d" % r).Value = "Info: block fuel CO2 > kernel sector energy CO2, %s (count of years)" % sec
            ws.Range("D%d:E%d" % (r, r)).Value = (s, "count")
            for c in DATA_COLS:
                ws.Cells(r, c).Formula = "=--(%s%s%d<-1E-9)" % (IP, col(c), rc + 2)
            ws.Range("F%d" % r).Formula = "=SUM(%s%d:%s%d)" % (L0, r, L1, r)
            r += 1
    ws.Range("A%d" % r).Value = "Info: double-count correction, scenario 2 (MtCO2e; kernel - CPAT-method IPPU)"
    ws.Range("D%d:E%d" % (r, r)).Value = (2, "value")
    for c in DATA_COLS:
        ws.Cells(r, c).Formula = "=%s%s%d" % (IP, col(c), b2["dc"])
    ws.Range("F%d" % r).Formula = maxabs(r)
    ws.Range("%s%d:%s%d" % (L0, r, L1, r)).NumberFormat = "0.000"
    r += 2
    ws.Range("A%d" % r).Value = (
        "Note: the CPAT-method row reproduces CPAT exactly only in scenario 1; in scenario 2 it applies CPAT's rule to "
        "the kernel's own industrial energy CO2 change (non-kernel industry sectors held at the CPAT baseline).")
    return r0 + 1


def update_scenarios(wb, rows):
    ws = wb.Worksheets("Scenarios")
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task G (%s): sheet IPPU_Industry - kernel IPPU egy.mit.ghg.iprk.tot.<s> (rows %d / %d) = block process "
        "(np + no) + fp flagged as IPPU + non-block remainder (Manual inputs E%d IppuOther NONE|CPAT); it replaces "
        "CPAT's proportional IPPU (egy.mit.ghg.iprc.tot.<s>); double-count correction egy.mit.ghg.iprdc.tot.<s>. "
        "Fuel-CO2 reconciliation by sector in sections C1/C2."
        % (VER, rows[("B", 1)]["kern"], rows[("B", 2)]["kern"], MAN_SW))


def update_settings(wb, chk, rows, diff_note, stats):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: Task G section (row %d)." % (VER, chk)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task G: new sheet IPPU_Industry. CPAT IPPU baseline by gas and industrial energy CO2 stored (CPAT "
        "Mitigation rows 7566-7570, 6590); CPAT-method IPPU (all IPPU x industrial energy CO2 growth, kernel change "
        "only) vs kernel IPPU = block process np+no (o145) + fp counted as IPPU (ammonia feedstock flag) + non-block "
        "remainder (CPAT baseline - block baseline; IppuOther NONE holds it, CPAT scales it). Double-count correction "
        "row iprdc. Reconciliation of block fuel CO2 with Task F sector energy CO2 by sector (report only). Manual "
        "inputs rows %d-%d (switch E%d, fp flags E%d:E%d). Check section row %d. %s. Regression vs v0.12: %s."
        % (MAN_TITLE, MAN_R0 + 7, MAN_SW, MAN_R0, MAN_R0 + 7, chk, stats, diff_note))
    ws.Range("C%d" % (last + 1)).WrapText = False


# ----------------------------------------------------------------------------------------------- regression
SNAP = ("Mitigation_Industry", "Emissions_Industry", "Rebate_Industry", "Fund_Industry", "Data_Prices", "Scenarios",
        "Data_EF", "Data_Energy", "MTInputs")


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
        snap[("Manual inputs", bd)] = wb.Worksheets("Manual inputs").Range("A1:AE82").Value
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return snap


def maxdiff(a, b):
    m, where = 0.0, None
    for i, (ra, rb) in enumerate(zip(a, b)):
        for j, (x, y) in enumerate(zip(ra, rb)):
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                if abs(x - y) > m:
                    m, where = abs(x - y), (i + 1, j + 1)
            elif x != y and not (n_empty(x) and n_empty(y)):
                if isinstance(x, str) and isinstance(y, str):
                    continue        # labels (e.g. Scenarios note column) - text rows are not regression values
                return float("inf"), (i + 1, j + 1, x, y)
    return m, where


def n_empty(v):
    return v in (None, "")


def verify(wb, snap, rows, chk):
    xl, st, ck = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Check")
    ip = wb.Worksheets("IPPU_Industry")
    max_rows = [r for r in range(1, ck.UsedRange.Rows.Count + 1) if str(ck.Cells(r, 1).Value).startswith("Max |")]
    legacy_only = {2, 3, 4, 584, 585, 1381}
    errs, stats = [], {}
    c30 = DATA_COLS[8]
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        bad = [r for r in max_rows if r not in legacy_only and
               (not isinstance(ck.Cells(r, 4).Value, (int, float)) or abs(ck.Cells(r, 4).Value) > 1e-9)]
        errs += ["%s: Check row %d %s = %s" % (bd, r, ck.Cells(r, 1).Value, ck.Cells(r, 4).Value) for r in bad]
        dmax = 0.0
        for n in SNAP + ("Check", "Manual inputs"):
            old = snap[(n, bd)]
            ws = wb.Worksheets(n)
            new = ws.Range("A1:%s%d" % ("AE" if n == "Manual inputs" else "AI", len(old))).Value
            d, w = maxdiff(old, new)
            if d > 1e-9:
                errs.append("%s: %s differs from v0.12 by %s at %s" % (bd, n, d, w))
            dmax = max(dmax, d)
        b2 = rows[("B", 2)]
        stats[bd] = tuple(ip.Cells(b2[k], c30).Value for k in ("cpat", "blk", "oth", "kern", "dc"))
        print("bundle %-6s failing Max: %d | max diff vs v0.12 %.3g | 2030 s2: CPAT-method %.2f block %.2f "
              "non-block %.2f kernel %.2f correction %.3f | Check G = %s"
              % ((bd, len(bad), dmax) + stats[bd] + (ck.Cells(chk, 4).Value,)))
        for ws in wb.Worksheets:
            vals = ws.UsedRange.Value
            if isinstance(vals, tuple):
                for i, rv in enumerate(vals):
                    for j, v in enumerate(rv):
                        if isinstance(v, int) and v < -2146820000:
                            errs.append("%s: error value %s!R%dC%d" % (bd, ws.Name, ws.UsedRange.Row + i,
                                                                       ws.UsedRange.Column + j))
    # switch test: IppuOther = CPAT on a priced bundle
    st.Range("B10").Value = "1A"
    mn = wb.Worksheets("Manual inputs")
    mn.Range("E%d" % MAN_SW).Value = "CPAT"
    xl.CalculateFull()
    b1, b2 = rows[("B", 1)], rows[("B", 2)]
    v = [ip.Cells(r, c30).Value for r in (b1["cpat"], b1["oth"], b2["cpat"], b2["oth"])]
    exp = v[1] * v[2] / v[0]
    print("IppuOther=CPAT (1A, 2030): non-block s2 %.4f expected %.4f | Check G = %s"
          % (v[3], exp, ck.Cells(chk, 4).Value))
    if abs(v[3] - exp) > 1e-9 or abs(ck.Cells(chk, 4).Value) > 1e-9:
        errs.append("IppuOther=CPAT switch test failed")
    mn.Range("E%d" % MAN_SW).Value = "NONE"
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    if errs:
        raise ValueError("verification failed:\n" + "\n".join(errs[:40]))
    return stats


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
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
        print("snapshot v0.12 ...")
        snap = snapshot(wb)
        build_manual(wb)
        rows = build_ippu(wb)
        chk = update_check(wb, rows)
        update_scenarios(wb, rows)
        xl.CalculateFull()
        print("verify ...")
        stats = verify(wb, snap, rows, chk)
        s3a = stats["3A"]
        txt = ("2030, bundle 3A scenario 2: CPAT-method IPPU %.1f, kernel IPPU %.1f MtCO2e (block %.1f, non-block "
               "%.1f), correction %.2f" % (s3a[0], s3a[3], s3a[1], s3a[2], s3a[4]))
        update_settings(wb, chk, rows, "all prior sheets identical for every bundle (max |diff| 0); new rows only",
                        txt)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        if os.environ.get("KEEP_DEBUG") and wb is not None:
            try:
                wb.SaveAs(os.path.join(HERE, "_debug_v0_13.xlsx"))
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
