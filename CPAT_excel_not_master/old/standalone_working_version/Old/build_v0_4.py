"""Build CPAT_Industry_Kernel_Egypt_v0.4.xlsx from v0.3 (Task E: Egypt price paths, policy bundles, stream interface).

Adds a 'Scenarios' sheet (EGY) with carbon price paths, the EgyptResultsInitial policy bundles (1A..3C plus LEGACY),
an active-bundle selector (Settings!B10) and the resolved scenario-2 series. Data_Prices scenario-2 policy rows
(cptraj, etstraj, nce) become lookups of those series; new interface rows (pptraj, encov, obrsh, obrrb, fundsh,
fundsp) are added for both scenarios. CBAM block rows o4-o6 read the interface rows; the fund-driven 'Additional ER'
rows now read o4 (shadow price) instead of o5 (fund size). Manual inputs E21 follows the bundle's process flag.

With Settings!B10 = LEGACY every value equals v0.3 (regression diff ~1e-15).

Run with Excel installed:  python build_v0_4.py
"""
import datetime
import os
import shutil

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.3.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.3.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.4.xlsx")

BASE_T0, CTAX_T0, NROWS = 245, 644, 156
BLOCKS = ((BASE_T0, "R5C10"), (CTAX_T0, "R404C10"))   # (first block row, scenario-number cell)
FUELS = ["coa", "nga", "gso", "die", "lpg", "ker", "oop", "bio", "ren"]
YEARS = list(range(2021, 2046))                        # K..AI
YC0 = 11                                               # column K
DATA_COLS = range(12, 32)                              # L..AE (2022-2041): CPAT data horizon

GREEN, BLUE, TAN = 0xDEF1EB, 0xF1E6DC, 0xC4D9DD       # EBF1DE, DCE6F1, DDD9C4 as BGR
TITLE, SECTION, REVIEW = 0x50B000, 0x50D092, 0xCCF2FF  # 00B050, 92D050, FFF2CC as BGR


def col(c):  # 1-based -> letter
    return chr(64 + c) if c <= 26 else "A" + chr(64 + c - 26)


def copy_formats(src, dst, tries=10):
    """Row format copy via clipboard; retries because other Excel instances can hold the clipboard."""
    import time
    for i in range(tries):
        try:
            src.Copy()
            dst.PasteSpecial(-4122)
            return
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(1)


def path(start, f):
    return [f(y) if y >= start else 0.0 for y in YEARS]


PATHS = [  # (id, description, source, values K..AI or None = legacy)
    ("LEGACY_CP", "CPAT default carbon tax path in the legacy file", "CPAT 1.0pre_456 Mitigation row 8582 (cptraj.2)", None),
    ("LEGACY_ETS", "CPAT default auctioned ETS path in the legacy file", "CPAT 1.0pre_456 Mitigation row 8583 (etstraj.2)", None),
    ("ZERO", "No price", "-", [0.0] * len(YEARS)),
    ("FLAT20_2028", "USD20/tCO2 flat from 2028", "EgyptResultsInitial 1A/2A/2B; CPAT runs EG1/EG2", path(2028, lambda y: 20.0)),
    ("RAMP_EG3", "USD5 in 2028, +7.5/yr (USD20 in 2030, USD50 in 2034, keeps rising)",
     "CPAT run EG3 (Outputs), behind the reported 3A-3C numbers", path(2028, lambda y: 5.0 + 7.5 * (y - 2028))),
    ("RAMP50_2034", "USD5 in 2028, +7.5/yr to USD50 in 2034, flat after",
     "EgyptResultsInitial text ('USD5 to USD50 by 2034')", path(2028, lambda y: min(50.0, 5.0 + 7.5 * (y - 2028)))),
]
REF_YEAR, REF_PRICE = 2030, 20.0
for _p in PATHS:
    if _p[3] is not None and _p[0] != "ZERO":
        assert _p[3][YEARS.index(REF_YEAR)] == REF_PRICE, _p[0]

BUNDLE_HDR = ("Bundle", "Name", "Energy CO2 price path", "ETS path", "Industry energy covered (share)",
              "Process emissions covered (0/1)", "Process price path", "OBR share (theta)", "Fund share (phi)",
              "Revenue use", "Energy scope in main CPAT (Task M)")
BUNDLES = [
    ("LEGACY", "CPAT default policy scenario (regression)", "LEGACY_CP", "LEGACY_ETS", 1, 0, "LEGACY_CP", 0, 0,
     "CPAT default", "All sectors"),
    ("1A", "Comprehensive levy + CBAM process fee", "FLAT20_2028", "ZERO", 1, 1, "FLAT20_2028", 0, 0,
     "Public investment (non-green)", "All sectors"),
    ("2A", "Upstream levy, public investment", "FLAT20_2028", "ZERO", 1, 0, "FLAT20_2028", 0, 0,
     "Public investment (non-green)", "All sectors"),
    ("2B", "Upstream levy, households", "FLAT20_2028", "ZERO", 1, 0, "FLAT20_2028", 0, 0,
     "Households (per capita)", "All sectors"),
    ("3A", "Downstream heavy industry, households", "RAMP_EG3", "ZERO", 1, 1, "RAMP_EG3", 0, 0,
     "Households (per capita)", "Industry only"),
    ("3B", "Downstream heavy industry, free allocation (benchmark)", "RAMP_EG3", "ZERO", 1, 1, "RAMP_EG3", 1, 0,
     "Firms: output-based free allocation", "Industry only"),
    ("3C", "Downstream heavy industry, abatement rebate", "RAMP_EG3", "ZERO", 1, 1, "RAMP_EG3", 0, 1,
     "Firms: abatement rebate fund", "Industry only"),
]

# interface series: (var, unit, description, writer, readers)
IFACE = [
    ("cptraj", "$/ton CO2 real", "Carbon price on energy CO2 (active bundle)", "Stream 2 (E)",
     "kernel via nce; CBAM block o2 (fuel and, until Task C, process)"),
    ("etstraj", "$/ton CO2 real", "Auctioned ETS price (active bundle)", "Stream 2 (E)", "CBAM block o3"),
    ("pptraj", "$/ton CO2 real", "Carbon price on CBAM process emissions = process path x process flag",
     "Stream 2 (E)", "Stream 1 (Task C: process pricing; L: CBAM obligations)"),
    ("encov", "share", "Share of industry energy CO2 covered by the energy price", "Stream 2 (E)", "nce (this sheet)"),
    ("obrsh", "share", "Output-based rebate share theta (free allocation / benchmark)", "Stream 2 (E, I)",
     "Stream 1 (H: output response; L: CBAM price deduction)"),
    ("obrrb", "$/ton CO2 real", "Output-based rebate = theta x cptraj (value returned per t at benchmark)",
     "Stream 2 (E, I)", "CBAM block o6; Stream 1 (H, L)"),
    ("fundsh", "share", "Share of revenue recycled to the industrial abatement fund phi", "Stream 2 (E, J)",
     "CBAM block o5"),
    ("fundsp", "$/ton CO2 real", "Shadow price of fund-financed abatement (acts on intensity only). "
     "Placeholder 0 until Task J", "Stream 2 (J)", "CBAM block o4 -> Additional ER rows o107-o114"),
]


def build_scenarios(wb, legacy_cp, legacy_ets, cpf):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Manual inputs"))
    ws.Name = "Scenarios"
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    for c, w in zip(range(1, 12), (3, 14, 34, 22, 14, 14, 26, 14, 12, 26, 22)):
        ws.Columns(c).ColumnWidth = w
    ws.Range("K:AI").ColumnWidth = 9
    ws.Range("A1:AI1").Interior.Color = TITLE
    ws.Range("B1").Value = "Scenarios (EGY): carbon price paths, policy bundles, active-bundle series"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 14
    ws.Range("B1").Font.Color = 0xFFFFFF
    ws.Range("B2").Value = ("Policy scenario 2 runs the bundle selected in Settings!B10 (one at a time). Section 4 is "
                            "mirrored into Data_Prices by code. LEGACY reproduces v0.3 / CPAT default values exactly.")
    ws.Range("B3").Value = "Change only green cells. Bundles from EgyptResultsInitial.docx Table 1/2."

    def band(r, text):
        ws.Range("A%d:AI%d" % (r, r)).Interior.Color = SECTION
        ws.Range("A%d" % r).Value = text.split(" ", 1)[0]
        ws.Range("B%d" % r).Value = text.split(" ", 1)[1]
        ws.Range("A%d:B%d" % (r, r)).Font.Bold = True

    # 1. price paths
    r = 5
    band(r, "1. Carbon price paths ($/ton CO2 real)")
    hdr = r + 1
    ws.Range("B%d:J%d" % (hdr, hdr)).Value = ("Path", "Description", "Source", "", "", "", "", "Unit", "Scenario")
    for k, y in enumerate(YEARS):
        ws.Cells(hdr, YC0 + k).Value = y
    ws.Range("A%d:AI%d" % (hdr, hdr)).Font.Bold = True
    p0 = hdr + 1
    for k, (pid, desc, src, vals) in enumerate(PATHS):
        rr = p0 + k
        if vals is None:
            vals = legacy_cp if pid == "LEGACY_CP" else legacy_ets
        ws.Range("B%d:D%d" % (rr, rr)).Value = (pid, desc, src)
        ws.Range("I%d:J%d" % (rr, rr)).Value = ("$/ton CO2 real", "EGY")
        ws.Range("K%d:AI%d" % (rr, rr)).Value = tuple(vals)
        ws.Range("K%d:AI%d" % (rr, rr)).Interior.Color = GREEN
    p1 = p0 + len(PATHS) - 1
    ws.Range("K%d:AI%d" % (p0, p1)).NumberFormat = "0.0"
    ws.Range("B%d" % (p1 + 1)).Value = ("Kernel and Data_Prices use 2022-2041 (L..AE), the CPAT data horizon. "
                                        "Reported EgyptResultsInitial metrics refer to 2030, when every non-legacy "
                                        "bundle prices USD20/tCO2. RAMP_EG3 (3A-3C default) and RAMP50_2034 differ "
                                        "only after 2034.")

    # 2. bundles
    r = p1 + 3
    band(r, "2. Policy bundles (one row per bundle)")
    bh = r + 1
    ws.Range("B%d:L%d" % (bh, bh)).Value = BUNDLE_HDR
    ws.Range("B%d:L%d" % (bh, bh)).Font.Bold = True
    ws.Range("B%d:L%d" % (bh, bh)).WrapText = True
    b0 = bh + 1
    for k, b in enumerate(BUNDLES):
        ws.Range("B%d:L%d" % (b0 + k, b0 + k)).Value = b
    b1 = b0 + len(BUNDLES) - 1
    ws.Range("C%d:L%d" % (b0, b1)).Interior.Color = GREEN
    for c in "DEH":  # path pick-lists
        v = ws.Range("%s%d:%s%d" % (c, b0, c, b1)).Validation
        v.Delete()
        v.Add(3, 1, 1, "=$B$%d:$B$%d" % (p0, p1))

    # 3. active bundle + pass-through factors
    r = b1 + 2
    band(r, "3. Active bundle (policy scenario 2) and carbon price pass-through factors")
    ws.Range("B%d:L%d" % (r + 1, r + 1)).Value = BUNDLE_HDR
    ws.Range("B%d:L%d" % (r + 1, r + 1)).Font.Bold = True
    act = r + 2
    ws.Range("B%d" % act).Formula = "=Settings!$B$10"
    ws.Range("B%d" % act).Interior.Color = TAN
    for c in range(3, 13):
        ws.Cells(act, c).Formula = "=INDEX(%s$%d:%s$%d,MATCH($B$%d,$B$%d:$B$%d,0))" % (
            col(c), b0, col(c), b1, act, b0, b1)
    A = {k: "$%s$%d" % (col(c), act) for k, c in
         (("cp", 4), ("ets", 5), ("encov", 6), ("proc", 7), ("pp", 8), ("theta", 9), ("phi", 10))}
    f0 = act + 2
    ws.Range("B%d:J%d" % (f0, f0)).Value = ("Fuel", "Pass-through factor: price units per $/ton CO2", "",
                                            "Factor", "", "", "", "Unit", "Source")
    ws.Range("B%d:J%d" % (f0, f0)).Font.Bold = True
    units = {"coa": "$/GJ", "nga": "$/GJ", "gso": "$/liter", "die": "$/liter", "lpg": "$/liter",
             "ker": "$/liter", "oop": "$/bbl", "bio": "$/GJ", "ren": "$/kwh"}
    cpf_row = {}
    for k, fu in enumerate(FUELS):
        rr = f0 + 1 + k
        cpf_row[fu] = rr
        ws.Range("B%d:C%d" % (rr, rr)).Value = (fu, "CO2 content per native unit (tCO2/unit)")
        ws.Range("E%d" % rr).Value = cpf[fu]
        ws.Range("E%d" % rr).Interior.Color = GREEN
        ws.Range("I%d:J%d" % (rr, rr)).Value = ("%s per $/tCO2" % units[fu],
                                                "Implied by CPAT: legacy nce.2 / cptraj.2 (exactly proportional)")

    # 4. resolved series
    r = f0 + len(FUELS) + 2
    band(r, "4. Active-bundle series for scenario 2 (read by Data_Prices via code in column G)")
    sh = r + 1
    ws.Range("B%d:J%d" % (sh, sh)).Value = ("Variable", "Description", "Written by", "Unit", "Source",
                                            "Input Code", "Read by", "", "Scenario")
    for k, y in enumerate(YEARS):
        ws.Cells(sh, YC0 + k).Value = y
    ws.Range("A%d:AI%d" % (sh, sh)).Font.Bold = True
    rows = {}
    rr = sh + 1
    for var, unit, desc, who, readers in IFACE:
        rows[var] = rr
        ws.Range("B%d:J%d" % (rr, rr)).Value = (var, desc, who, unit, "Scenarios", "egy.mit.%s.2" % var,
                                                readers, "", "EGY")
        rr += 1
    for fu in FUELS:
        rows["nce." + fu] = rr
        ws.Range("B%d:J%d" % (rr, rr)).Value = ("nce", "New-policy wedge, %s = cptraj x encov x factor" % fu,
                                                "Stream 2 (E)", units[fu], "Scenarios",
                                                "egy.mit.nce.ind.%s.a.2" % fu, "kernel atp rows", "", "EGY")
        rr += 1
    s1 = rr - 1
    pr = "INDEX(K$%d:K$%d,MATCH(%%s,$B$%d:$B$%d,0))" % (p0, p1, p0, p1)
    F = {
        "cptraj": "=" + pr % A["cp"],
        "etstraj": "=" + pr % A["ets"],
        "pptraj": "=" + pr % A["pp"] + "*" + A["proc"],
        "encov": "=" + A["encov"],
        "obrsh": "=" + A["theta"],
        "obrrb": "=K%d*K%d" % (rows["obrsh"], rows["cptraj"]),
        "fundsh": "=" + A["phi"],
        "fundsp": None,
    }
    for var, f in F.items():
        rng = ws.Range("K%d:AI%d" % (rows[var], rows[var]))
        if f is None:
            rng.Value = 0
            rng.Interior.Color = GREEN
        else:
            rng.Formula = f
    for fu in FUELS:
        ws.Range("K%d:AI%d" % (rows["nce." + fu], rows["nce." + fu])).Formula = "=K%d*K%d*$E$%d" % (
            rows["cptraj"], rows["encov"], cpf_row[fu])
    ws.Range("G%d:G%d" % (sh + 1, s1)).Interior.Color = BLUE
    ws.Range("K%d:AI%d" % (sh + 1, s1)).NumberFormat = "0.000"

    # 5. interface contract
    r = s1 + 2
    band(r, "5. Interface contract between work streams (merge rules)")
    lines = [
        "Stream 1 = product/CBAM side (tasks A, B, C, D, H, K, L): owns Manual inputs rows 18-47 and CBAM block rows o7-o155.",
        "Stream 2 = energy/policy side (tasks E, F, I, J): owns Scenarios, Data_Prices policy rows, kernel sector blocks, block rows o2-o6.",
        "Interface = the codes in section 4 (Data_Prices egy.mit.<var>.<s>) plus block rows o2-o6. Stream 2 writes, stream 1 reads.",
        "o2 Carbon Tax = cptraj; o3 ETS = etstraj; o4 Total Shadow Price = fundsp ($/tCO2, intensity only); "
        "o5 Revenues Fund = fundsh x Total Revenues (o155); o6 Output-based rebate = obrrb.",
        "Price seen by: efficiency / intensity (fuel and process ER) = full price (+ o4); output / usage = price - obrrb (Task H, I).",
        "Process emissions price = pptraj (Task C switches block process rows from o2 to pptraj; equal for all current bundles).",
        "Manual inputs E21 (process emissions priced) = active bundle 'Process emissions covered'; per-sector detail stays in Manual inputs (Task C).",
        "Rules: find rows by code/label, never by row number; never insert rows in the other stream's area; each stream adds its own Check section; "
        "merge = run both stream build scripts on v0.4.",
    ]
    for k, t in enumerate(lines):
        ws.Range("B%d" % (r + 1 + k)).Value = t
    ws.Range("A%d:A%d" % (5, r + len(lines))).Font.Bold = True
    ws.Range("A1").Select()
    return {"act": act, "proc": A["proc"], "rows": rows, "p0": p0, "b0": b0, "b1": b1}


def update_data_prices(wb):
    ws = wb.Worksheets("Data_Prices")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    codes = {ws.Cells(r, 1).Value: r for r in range(2, last + 1)}
    look = "=INDEX(Scenarios!C,MATCH(RC1,Scenarios!C7,0))"
    for code in ["egy.mit.cptraj.2", "egy.mit.etstraj.2"] + ["egy.mit.nce.ind.%s.a.2" % f for f in FUELS]:
        r = codes[code]
        for c in DATA_COLS:
            ws.Cells(r, c).FormulaR1C1 = look
        ws.Range("G%d" % r).Value = str(ws.Range("G%d" % r).Value) + " - active bundle (Scenarios)"
        ws.Range("H%d" % r).Value = "Scenarios (legacy %d)" % int(ws.Range("H%d" % r).Value)
        ws.Range("J%d" % r).Value = "EGY"
    first = last + 1
    new = [(v, s) for s in (1, 2) for v, *_ in IFACE if v not in ("cptraj", "etstraj")]
    copy_formats(ws.Rows(last), ws.Rows("%d:%d" % (first, first + len(new) - 1)))
    wb.Application.CutCopyMode = False
    meta = {v: (u, d) for v, u, d, *_ in IFACE}
    for k, (var, s) in enumerate(new):
        r = first + k
        code = "egy.mit.%s.%d" % (var, s)
        u, d = meta[var]
        src = "Scenarios" if s == 2 else "- (baseline: no new policy)"
        ws.Range("A%d:J%d" % (r, r)).Value = (code, "-", "all", var, s, u, d, src, code, "EGY")
        for c in DATA_COLS:
            if s == 2:
                ws.Cells(r, c).FormulaR1C1 = look
            else:
                ws.Cells(r, c).Value = 0
    return first, first + len(new) - 1


def update_blocks(wb):
    ws = wb.Worksheets("Mitigation_Industry")
    for t0, scen in BLOCKS:
        for off, label in ((2, "Carbon Tax"), (3, "ETS Price (Auctioned)"), (4, "Total Shadow Price"), (5, "Revenues Fund")):
            if ws.Range("D%d" % (t0 + off)).Value != label:
                raise ValueError("block row o%d not found at %d" % (off, t0 + off))
        if any(ws.Cells(t0 + 6, c).Formula not in ("", None) for c in range(1, 36)):
            raise ValueError("block row o6 not empty")
        lk = "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))"
        o4, o5, o6 = t0 + 4, t0 + 5, t0 + 6
        ws.Range("D%d:I%d" % (o4, o4)).Value = ("Total Shadow Price", "$/ton CO2", "Data_Prices", "", "",
                                                "Fund-financed abatement shadow price (Task J); intensity only")
        ws.Range("G%d" % o4).FormulaR1C1 = '=Settings!R3C2&".mit.fundsp."&%s' % scen
        ws.Range("D%d:I%d" % (o5, o5)).Value = ("Revenues Fund", "units of o155", "Data_Prices", "", "",
                                                "fundsh x Total Revenues (o155); converted to o4 in Task J")
        ws.Range("G%d" % o5).FormulaR1C1 = '=Settings!R3C2&".mit.fundsh."&%s' % scen
        ws.Range("D%d:I%d" % (o6, o6)).Value = ("Output-based rebate", "$/ton CO2 real", "Data_Prices", "", "",
                                                "theta x carbon tax; deducted from the price seen by output (Tasks H, I)")
        ws.Range("G%d" % o6).FormulaR1C1 = '=Settings!R3C2&".mit.obrrb."&%s' % scen
        ws.Range("K%d:AE%d" % (o4, o6)).ClearContents()
        ws.Range("K%d:AE%d" % (o4, o6)).Interior.Pattern = -4142
        for c in DATA_COLS:
            ws.Cells(o4, c).FormulaR1C1 = lk
            ws.Cells(o6, c).FormulaR1C1 = lk
            ws.Cells(o5, c).FormulaR1C1 = "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))*R[150]C"
        ws.Range("G%d:G%d" % (o4, o6)).Interior.Color = BLUE
        # fund-driven Additional ER rows: price = o4 (shadow price), not o5 (fund size)
        for off in range(107, 115):
            r = t0 + off
            old, new = "R[-%d]C" % (off - 5), "R[-%d]C" % (off - 4)
            for c in DATA_COLS:
                f = ws.Cells(r, c).FormulaR1C1
                if old not in f:
                    raise ValueError("unexpected Additional ER formula at %d: %s" % (r, f))
                ws.Cells(r, c).FormulaR1C1 = f.replace(old, new)
            if old in str(ws.Cells(r, 11).FormulaR1C1):
                ws.Cells(r, 11).FormulaR1C1 = ws.Cells(r, 11).FormulaR1C1.replace(old, new)


def update_manual_inputs(wb, sc):
    ws = wb.Worksheets("Manual inputs")
    ws.Range("E21").Formula = "=Scenarios!%s=1" % sc["proc"]
    ws.Range("E21").Interior.Color = TAN
    ws.Range("D21").Value = ("Include non-fuel process emissions in ETS/CT? (legacy Dashboard!G20; v0.4: set by the "
                             "active bundle, Scenarios)")


def update_settings(wb, sc):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype v0.4"
    ws.Range("A10").Value = "Policy bundle for scenario 2 (EGY)"
    ws.Range("B10").Value = "LEGACY"
    ws.Range("B10").Interior.Color = GREEN
    v = ws.Range("B10").Validation
    v.Delete()
    v.Add(3, 1, 1, "=Scenarios!$B$%d:$B$%d" % (sc["b0"], sc["b1"]))
    ws.Range("C10").Value = ("Bundle run as policy scenario 2 (Scenarios section 2). LEGACY = CPAT default "
                             "(regression vs v0.3/legacy). Check-sheet legacy comparisons are valid for LEGACY only.")
    if ws.Range("A17").Value != "Manual inputs":
        raise ValueError("Settings layout changed")
    ws.Rows(18).Insert()
    ws.Range("A18").Value = "Scenarios"
    ws.Range("C18").Value = ("EGY: carbon price paths, policy bundles 1A-3C (+LEGACY), active-bundle series mirrored "
                             "into Data_Prices by code, pass-through factors, interface contract between work streams.")
    if ws.Range("A22").Value != "Data_Prices":
        raise ValueError("Settings layout changed")
    ws.Range("C22").Value = str(ws.Range("C22").Value) + (" v0.4: scenario-2 cptraj/etstraj/nce read Scenarios; "
                                                         "interface rows pptraj, encov, obrsh, obrrb, fundsh, fundsp.")
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " v0.4: plus Scenarios regression section."
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = "v0.4"
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task E: Scenarios sheet (price paths LEGACY/ZERO/FLAT20_2028/RAMP_EG3/RAMP50_2034; bundles LEGACY,1A,2A,2B,3A,3B,3C; "
        "selector Settings!B10). Scenario-2 cptraj/etstraj/nce now = active-bundle series (nce = cptraj x encov x factor). "
        "New interface rows pptraj/encov/obrsh/obrrb/fundsh/fundsp (both scenarios). CBAM block o4 = fundsp, o5 = fundsh x o155, "
        "o6 = obrrb; Additional ER rows o107-o114 read o4 instead of o5. Manual inputs E21 = bundle process flag. "
        "LEGACY: regression diff vs v0.3 ~1e-15 (floating point of cptraj x factor).")
    ws.Range("C%d" % (last + 1)).WrapText = False


def build_check(wb, legacy):
    ws = wb.Worksheets("Check")
    r0 = ws.Cells(ws.Rows.Count, 1).End(-4162).Row + 3
    ws.Range("A%d" % r0).Value = ("Scenario-2 policy inputs (Data_Prices, now from Scenarios) vs values stored in v0.3 / "
                                  "CPAT_1_0pre_456; expected same when Settings!B10 = LEGACY")
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Code", "", "", "Scenario", "Series", "max |diff|", "", "", "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    r = hdr + 1
    for code, vals in legacy.items():
        for k, series in enumerate(("v0.3", "v0.4", "diff")):
            ws.Range("A%d" % (r + k)).Value = code
            ws.Range("D%d:E%d" % (r + k, r + k)).Value = (2, series)
        for c in DATA_COLS:
            ws.Cells(r, c).Value = vals[c - 1]
        L = col(DATA_COLS[0])
        ws.Range("%s%d:AE%d" % (L, r + 1, r + 1)).Formula = "=INDEX(Data_Prices!%s:%s,MATCH($A%d,Data_Prices!$A:$A,0))" % (L, L, r + 1)
        ws.Range("%s%d:AE%d" % (L, r + 2, r + 2)).Formula = "=%s%d-%s%d" % (L, r + 1, L, r)
        ws.Range("F%d" % (r + 2)).Formula = "=MAX(MAX(L%d:AE%d),-MIN(L%d:AE%d))" % (r + 2, r + 2, r + 2, r + 2)
        ws.Range("I%d" % (r + 2)).Value = "same if LEGACY"
        r += 4
    ws.Range("A%d" % (r0 + 1)).Value = "Max |v0.4 - v0.3| scenario-2 price inputs (valid when Settings!B10 = LEGACY)"
    ws.Range("D%d" % (r0 + 1)).Formula = '=MAX(F%d:F%d)' % (hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW


def read_legacy_policy(wb):
    ws = wb.Worksheets("Data_Prices")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    vals = {row[0]: row for row in ws.Range("A1:AI%d" % last).Value}
    legacy = {c: vals[c] for c in ["egy.mit.cptraj.2", "egy.mit.etstraj.2"] + ["egy.mit.nce.ind.%s.a.2" % f for f in FUELS]}
    cp = legacy["egy.mit.cptraj.2"]
    cpf = {}
    for fu in FUELS:
        n = legacy["egy.mit.nce.ind.%s.a.2" % fu]
        f = n[30] / cp[30]  # 2041 (largest price)
        for c in DATA_COLS:
            x, p = n[c - 1] or 0.0, cp[c - 1] or 0.0
            if abs(x - f * p) > 1e-12 * max(1.0, abs(x)):
                raise ValueError("nce.%s not proportional to cptraj in column %d" % (fu, c))
        cpf[fu] = f
    lcp = [cp[c - 1] for c in range(YC0, YC0 + len(YEARS))]
    lets = [legacy["egy.mit.etstraj.2"][c - 1] for c in range(YC0, YC0 + len(YEARS))]
    return legacy, cpf, lcp, lets


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
        legacy, cpf, lcp, lets = read_legacy_policy(wb)
        sc = build_scenarios(wb, lcp, lets, cpf)
        update_settings(wb, sc)
        update_data_prices(wb)
        update_blocks(wb)
        update_manual_inputs(wb, sc)
        build_check(wb, legacy)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        xl.DisplayAlerts = False
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
