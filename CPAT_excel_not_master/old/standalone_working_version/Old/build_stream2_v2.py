"""Build CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v2.xlsx from ..._Stream2_v1 (Task I: output-based rebating).

Stream 2 branch (v0.7 -> Stream2_v1 Task F -> this). Merge-friendly: one NEW sheet, appended Check / Settings-log /
Scenarios rows, and ONE controlled formula edit in Mitigation_Industry (the usage-term price reference of the ener
rows, asserted against the v1 R1C1 pattern before it is rewritten). No CBAM-block rows, no Manual-inputs rows touched.

Mechanism (EgyptTaskReference.md section 3)
  Rebate per t of product k      obrpt_k = MIN(obrrb x EU benchmark_k, full carbon cost per t_k)
                                 obrrb   = theta x cptraj (Data_Prices egy.mit.obrrb.<s> = CBAM block o6)
                                 full carbon cost per t_k = block fuel price increase (o54-o61) + process price
                                 increase (o65-o72), $/t. The MIN caps the rebate at the carbon cost actually paid, so
                                 the net price never falls below the pre-tax price (urea: benchmark > intensity).
  Net carbon cost per t_k        ccnet_k = full cost - obrpt_k;   ppin_k = ccnet_k / product price   <- Task H hook
  Sector rebate share            rsh_j   = MIN(1, sum_k Q_k obrpt_k / sum_k Q_k fullcost_k)   (0 if no carbon cost)
  Net fuel price (usage term)    atpn_{j,f} = atp - (atp - ptp) x rsh_j      (bit-identical to atp when theta = 0)
  Kernel                         the ener usage term IF(atp(t-1)<=0,1,atp(t)/atp(t-1))^eU reads atpn instead of atp;
                                 the efficiency term keeps shp (= atp + shpw x ssc, full price); the CBAM block
                                 process-ER rows (o118-o125) keep the full process price -> both unchanged.
  Rebate cost / net revenue      obr_k = obrpt_k x Q_k / 1000 (USD million, unit of revf/revp/rev.cbam.tot);
                                 revnet.cbam.tot = rev.cbam.tot - obr.cbam.tot.
Why explicit net-price rows instead of the shp/ssc wedge: shp/ssc are separable per-fuel (shp.<fuel>.ind.<s>) x
per-sector (ssc.<sector>.<s>) series reserved for the Task J fund shadow price; carrying the rebate through them
would require rewriting every atp row (so that atp = net and shp = full) and would entangle Tasks I and J. The
net-price rows leave atp, shp, Data_Prices and the CBAM block untouched and are the natural hook for Task H.
Production is exogenous until Task H (Stream 1): the "output / usage" effect of the net price acts only through the
kernel energy usage term; the ppin rows are what Task H wires into output.

Added
  Rebate_Industry  New sheet (after Emissions_Industry), Mitigation_Industry layout (A-AI, years K..AI, data L..AE).
                   Per scenario: policy rows (obrsh, obrrb); 8-product blocks obrbm (benchmark), ccpt (full cost/t),
                   obrpu (rebate/t uncapped), obrpt (rebate/t used), ccnet (net cost/t), ppin (net price increase
                   share, Task H hook), obr (rebate cost, + total obr.cbam.tot); sector blocks irn/cem/nfm/mch: obr,
                   ccf (cost base), rsh, 9 atpn rows; revenue: rev (link o155), obr, revnet.cbam.tot.
                   Codes egy.mit.<var>.<prod>.<s>, egy.mit.<var>.ind.<sector>.<s>, egy.mit.atpn.<sector>.<fuel>.e.<s>.
  Check            Section 'Task I (Stream2 v2)': rebate = 0 when obrsh = 0; net cost in [0, full cost]; atpn in
                   [ptp, atp]; rsh in [0,1]; totals; revnet = rev - obr; efficiency-term price (sum of shp per
                   sector) unchanged vs v1 for the active bundle; ener totals identical to v1 for every theta = 0
                   bundle (v1 values stored per bundle, selected by Settings!B10); ener >= v1 for theta > 0 bundles.
  Settings         Title, C25 note, version-log row 'v0.7branch_Stream2_v2'.
  Scenarios        Note line 'Task I (stream 2, v0.7branch_Stream2_v2): ...'; 'Read by' of obrsh/obrrb extended.

Run with Excel installed (from this folder):  python build_stream2_v2.py
The script loops Settings!B10 over all bundles (before the changes, to store v1 values; after, to verify the Check
section is 0 for each) and saves with Settings!B10 = LEGACY.
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLUE, DATA_COLS, FUELS, GREEN, REVIEW, col, copy_formats

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v1.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v2.xlsx")
VER = "v0.7branch_Stream2_v2"
MI, RI = "Mitigation_Industry!", "Rebate_Industry!"

SECTORS = ("irn", "cem", "nfm", "mch")
HEADERS = {1: (5, 65, 125, 185), 2: (404, 464, 524, 584)}
PTP0, ATP0, SHP0, ENER0, TOT = 2, 12, 22, 49, 58       # row offsets from the sector header
BLOCK = {1: 245, 2: 644}                                 # CBAM block header rows; block row o(n) = header + n
O_OBRRB, O_BENCH, O_PROD, O_PRICE, O_FPI, O_PPI, O_REV = 6, 19, 31, 43, 54, 65, 155
Z_BENCH = 26                                             # block column Z (rows o19-o26): EU benchmark tCO2/t
PRODUCTS = ("stl.drg", "stl.scr", "stl.bof", "cmt.dry", "frt.amc", "frt.ure", "frt.nit", "alu.prp")
PROD_NAME = ("DRI-EAF steel", "Scrap-EAF steel", "BOF steel", "Clinker / cement", "Ammonia", "Urea",
             "Ammonium nitrate", "Primary aluminium")
PROD_SECTOR = ("irn", "irn", "irn", "cem", "mch", "mch", "mch", "nfm")
EU_BENCH = (0.481, 0.072, 1.37, 0.666, 1.484, 0.902, 0.302, 1.423)
BUNDLES = ("LEGACY", "1A", "2A", "2B", "3A", "3B", "3C")
L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
NF, NP = len(FUELS), len(PRODUCTS)

ENER_TEMPLATE = ("=RC[-1]*(1+R{g}C)^(R[-10]C5*R{d}C)*(1+R{p}C)*IF(R[-37]C[-1]<=0,1,R[-37]C/R[-37]C[-1])^R[-10]C6"
                 "*(IF(R[-27]C[-1]<=0,1,R[-27]C/R[-27]C[-1])^R[-10]C7/(1+R[-10]C8+R[-10]C9))^(1+R[-10]C6)")
USAGE_OLD = "IF(R[-37]C[-1]<=0,1,R[-37]C/R[-37]C[-1])"
USAGE_NEW = "IF(Rebate_Industry!R{n}C[-1]<=0,1,Rebate_Industry!R{n}C/Rebate_Industry!R{n}C[-1])"


def ener_template(h):
    return ENER_TEMPLATE.format(g=h + 34, d=h + 35, p=h + 36)


def rng(r):
    return "%s%d:%s%d" % (L0, r, L1, r)


# ----------------------------------------------------------------------------------------------- preconditions
def check_layout(wb):
    if "Rebate_Industry" in [ws.Name for ws in wb.Worksheets]:
        raise ValueError("sheet Rebate_Industry already exists")
    mi = wb.Worksheets("Mitigation_Industry")
    for s, hdrs in HEADERS.items():
        for h, sec in zip(hdrs, SECTORS):
            if mi.Cells(h, 3).Value != sec or mi.Cells(h, 4).Value != "Fuel Type" or mi.Cells(h, 10).Value != s:
                raise ValueError("sector header moved: row %d expected %s scenario %d" % (h, sec, s))
            for i, f in enumerate(FUELS):
                for off, var in ((PTP0, "ptp"), (ATP0, "atp"), (SHP0, "shp"), (ENER0, "ener")):
                    r = h + off + i
                    if mi.Cells(r, 2).Value != f or mi.Cells(r, 6).Value != var:
                        raise ValueError("%s row moved: row %d expected %s" % (var, r, f))
                r = h + ENER0 + i
                for c in DATA_COLS[1:]:
                    if mi.Cells(r, c).FormulaR1C1 != ener_template(h):
                        raise ValueError("ener formula at %s%d does not match the v1 pattern" % (col(c), r))
            if mi.Cells(h + TOT, 2).Value != "tot" or mi.Cells(h + TOT, 6).Value != "ener":
                raise ValueError("ener total row moved: row %d" % (h + TOT))
    for s, b in BLOCK.items():
        if mi.Cells(b, 10).Value != s or mi.Cells(b + O_OBRRB, 7).Value != "egy.mit.obrrb.%d" % s:
            raise ValueError("block o6 (obrrb) moved, scenario %d" % s)
        if (mi.Cells(b + O_REV, 8).Value != "egy.mit.rev.cbam.tot.%d" % s
                or mi.Cells(b + O_REV, 5).Value != "USD million"):
            raise ValueError("block o155 (total revenues) moved, scenario %d" % s)
        for i, p in enumerate(PRODUCTS):
            if mi.Cells(b + O_PRICE + i, 6).Value != p:
                raise ValueError("block product price row o%d moved (%s)" % (O_PRICE + i, p))
            if abs((mi.Cells(b + O_BENCH + i, Z_BENCH).Value or 0) - EU_BENCH[i]) > 1e-12:
                raise ValueError("block EU benchmark Z%d changed" % (b + O_BENCH + i))
            if not str(mi.Cells(b + O_PROD + i, 5).Value).startswith("kt"):
                raise ValueError("block production row o%d moved" % (O_PROD + i))
            if (mi.Cells(b + O_FPI + i, 4).Value != mi.Cells(b + O_PPI + i, 4).Value
                    or mi.Cells(b + O_FPI + i, 4).Value is None):
                raise ValueError("block price-increase rows o%d/o%d moved" % (O_FPI + i, O_PPI + i))
    st = wb.Worksheets("Settings")
    if st.Range("B10").Value != "LEGACY":
        raise ValueError("source must be saved with Settings!B10 = LEGACY")
    if st.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    sc = wb.Worksheets("Scenarios")
    if ([sc.Cells(r, 2).Value for r in range(17, 24)] != list(BUNDLES) or sc.Cells(16, 9).Value != "OBR share (theta)"
            or sc.Cells(26, 9).Value != "OBR share (theta)"):
        raise ValueError("Scenarios bundle table moved")


# ----------------------------------------------------------------------------------------------- v1 reference values
def collect_v1(wb):
    """Per bundle: scenario-2 ener totals and sum of shp per sector; plus (bundle-independent) scenario-1 ener totals."""
    xl, st, mi = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Mitigation_Industry")
    ref = {"theta": {}}
    for b in BUNDLES:
        st.Range("B10").Value = b
        xl.CalculateFull()
        ref["theta"][b] = wb.Worksheets("Scenarios").Range("I27").Value
        for h, sec in zip(HEADERS[2], SECTORS):
            ref[("ener", 2, sec, b)] = list(mi.Range(rng(h + TOT)).Value[0])
            shp = mi.Range("%s%d:%s%d" % (L0, h + SHP0, L1, h + SHP0 + NF - 1)).Value
            ref[("shp", 2, sec, b)] = [sum(row[c] for row in shp) for c in range(len(DATA_COLS))]
    for h, sec in zip(HEADERS[1], SECTORS):
        ref[("ener", 1, sec)] = list(mi.Range(rng(h + TOT)).Value[0])
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return ref


# ----------------------------------------------------------------------------------------------- Rebate_Industry
def build_rebate(wb):
    mi = wb.Worksheets("Mitigation_Industry")
    ws = wb.Worksheets.Add(After=wb.Worksheets("Emissions_Industry"))
    ws.Name = "Rebate_Industry"
    ws.Cells.Font.Name = "Arial"
    mi.Rows("1:2").Copy(ws.Rows(1))
    wb.Application.CutCopyMode = False
    for c in range(1, 36):
        ws.Columns(c).ColumnWidth = mi.Columns(c).ColumnWidth
    ws.Columns(4).ColumnWidth = 40
    ws.Columns(9).ColumnWidth = 55
    ws.Range("B2").Value = ("Mitigation Module - Egypt: Industry kernel - output-based rebating (Task I, Stream 2): "
                            "rebate per t = MIN(obrrb x EU benchmark, carbon cost per t); net fuel price atpn = atp - "
                            "(atp - ptp) x sector rebate share feeds the Mitigation_Industry ener USAGE term; shp "
                            "(efficiency term) and the block process ER keep the full price")
    out = {}

    def header(r, text, s, src_row):
        copy_formats(mi.Rows(src_row), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d:J%d" % (r, r)).Value = ("", text, "Item", "Unit", "Variable Code", "Input Code", "Output Code",
                                              "Note", s)
        ws.Range("K%d:AI%d" % (r, r)).Value = mi.Range("K%d:AI%d" % (src_row, src_row)).Value

    def line(r, key, sec, item, unit, var, inp, outcode, note, formula, src_row, fmt="0.0000"):
        copy_formats(mi.Rows(src_row), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d:F%d" % (r, r)).Value = (key, sec, item, unit, var)
        ws.Range("G%d" % r).Interior.ColorIndex = -4142
        ws.Range("H%d" % r).Interior.ColorIndex = -4142
        if inp:
            ws.Range("G%d" % r).Value = inp
            ws.Range("G%d" % r).Interior.Color = BLUE
        if outcode:
            ws.Range("H%d" % r).Formula = '=Settings!$B$3&".mit.%s"' % outcode
            ws.Range("H%d" % r).Interior.Color = BLUE
        ws.Range("I%d" % r).Value = note
        ws.Range(rng(r)).FormulaR1C1 = formula
        ws.Range(rng(r)).NumberFormat = fmt

    dp = "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))"
    r = 3
    for s in (1, 2):
        b = BLOCK[s]
        h0 = HEADERS[s][0]
        copy_formats(mi.Rows(3), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d" % r).Value = ("Output-based rebate - %s (scenario %d); CBAM block rows %d-%d, kernel rows %d-%d"
                                     % ("Baseline" if s == 1 else "Policy", s, b, b + O_REV, h0, HEADERS[s][-1] + TOT))
        r += 2
        header(r, "policy", s, h0)
        out[("first", s)] = r
        r_sh, r_rb = r + 1, r + 2
        line(r_sh, "obrsh", "", "Output-based rebate share theta (active bundle)", "share", "obrsh",
             "egy.mit.obrsh.%d" % s, None, "Data_Prices (Scenarios section 4)", dp, h0 + PTP0)
        line(r_rb, "obrrb", "", "Output-based rebate per t at benchmark = theta x cptraj", "$/ton CO2 real", "obrrb",
             "egy.mit.obrrb.%d" % s, None, "Data_Prices = CBAM block o6 (row %d); ETS not included (Task E)"
             % (b + O_OBRRB), dp, h0 + PTP0)
        out[("obrsh", s)] = r_sh
        r = r_rb + 2
        prod = {}
        blocks = (
            ("obrbm", "EU product benchmark", "tCO2/t",
             "CBAM block column Z rows o%d-o%d (Manual inputs Z30:Z37)" % (O_BENCH, O_BENCH + NP - 1),
             lambda i, p: "=Mitigation_Industry!R%dC%d" % (b + O_BENCH + i, Z_BENCH)),
            ("ccpt", "Full carbon cost per t = fuel + process price increase", "$/t",
             "CBAM block o%d + o%d" % (O_FPI, O_PPI),
             lambda i, p: "=Mitigation_Industry!R%dC+Mitigation_Industry!R%dC" % (b + O_FPI + i, b + O_PPI + i)),
            ("obrpu", "Rebate per t uncapped = obrrb x benchmark", "$/t", "memo",
             lambda i, p: "=R%dC*R%dC" % (r_rb, p["obrbm"])),
            ("obrpt", "Rebate per t used = MIN(uncapped, full carbon cost)", "$/t",
             "cap keeps the net price >= pre-tax price (binding when benchmark > priced intensity, e.g. urea)",
             lambda i, p: "=MIN(R%dC,R%dC)" % (p["obrpu"], p["ccpt"])),
            ("ccnet", "Net carbon cost per t = full cost - rebate", "$/t", "",
             lambda i, p: "=R%dC-R%dC" % (p["ccpt"], p["obrpt"])),
            ("ppin", "Net product price increase share (TASK H HOOK: wire into output)", "share",
             "net cost / product price (CBAM block o%d, USD/t)" % O_PRICE,
             lambda i, p: "=IF(Mitigation_Industry!R%dC<=0,0,R%dC/Mitigation_Industry!R%dC)"
             % (b + O_PRICE + i, p["ccnet"], b + O_PRICE + i)),
            ("obr", "Rebate cost = rebate per t x production / 1000", "USD million",
             "production kt: CBAM block o%d-o%d" % (O_PROD, O_PROD + NP - 1),
             lambda i, p: "=R%dC*Mitigation_Industry!R%dC/1000" % (p["obrpt"], b + O_PROD + i)),
        )
        for var, item, unit, note, fml in blocks:
            header(r, var, s, h0)
            for i, p in enumerate(PRODUCTS):
                rr = r + 1 + i
                prod[(var, i)] = rr
                refs = {v: prod.get((v, i), 0) for v in ("obrbm", "ccpt", "obrpu", "obrpt", "ccnet")}
                line(rr, p, PROD_SECTOR[i], "%s: %s" % (PROD_NAME[i], item), unit, var, None,
                     "%s.%s.%d" % (var, p, s), note if i == 0 else "", fml(i, refs), h0 + PTP0)
            r += 1 + NP
            if var == "obr":
                line(r, "tot", "", "Total rebate cost, CBAM products", unit, var, None, "obr.cbam.tot.%d" % s, "",
                     "=SUM(R[-%d]C:R[-1]C)" % NP, h0 + TOT)
                out[("obr", s)] = r
                r += 1
            r += 1
        for h, sec in zip(HEADERS[s], SECTORS):
            header(r, sec, s, h)
            idx = [i for i, ps in enumerate(PROD_SECTOR) if ps == sec]
            r_ob, r_cf, r_rs = r + 1, r + 2, r + 3
            line(r_ob, "tot", sec, "Sector rebate = sum of product rebate costs", "USD million", "obr", None,
                 "obr.ind.%s.%d" % (sec, s), "products: " + ", ".join(PRODUCTS[i] for i in idx),
                 "=" + "+".join("R%dC" % prod[("obr", i)] for i in idx), h + TOT)
            line(r_cf, "tot", sec, "Sector carbon-cost base = sum of full cost per t x production / 1000", "USD million",
                 "ccf", None, "ccf.ind.%s.%d" % (sec, s), "",
                 "=" + "+".join("R%dC*Mitigation_Industry!R%dC/1000" % (prod[("ccpt", i)], b + O_PROD + i)
                                for i in idx), h + TOT)
            line(r_rs, "tot", sec, "Sector rebate share rsh = MIN(1, rebate / carbon-cost base)", "share", "rsh", None,
                 "rsh.ind.%s.%d" % (sec, s), "applied to the kernel fuel carbon wedge (atp - ptp)",
                 "=IF(R%dC<=0,0,MIN(1,R%dC/R%dC))" % (r_cf, r_ob, r_cf), h + TOT)
            out[("obr", s, sec)], out[("rsh", s, sec)] = r_ob, r_rs
            for i, f in enumerate(FUELS):
                rr, src = r_rs + 1 + i, h + ATP0 + i
                line(rr, f, sec, "%s: net post-tax price = atp - (atp - ptp) x rsh" % mi.Cells(src, 4).Value,
                     mi.Cells(src, 5).Value, "atpn", None, "atpn.%s.%s.e.%d" % (sec, f, s),
                     "read by the ener usage term, Mitigation_Industry row %d" % (h + ENER0 + i),
                     "=Mitigation_Industry!R%dC-(Mitigation_Industry!R%dC-Mitigation_Industry!R%dC)*R%dC"
                     % (src, src, h + PTP0 + i, r_rs), src)
                out[("atpn", s, sec, f)] = rr
            r = r_rs + NF + 2
        header(r, "revenue", s, h0)
        line(r + 1, "tot", "", "Total revenues (CBAM block o155)", "USD million", "rev", None, None,
             "link to egy.mit.rev.cbam.tot.%d (Mitigation_Industry row %d)" % (s, b + O_REV),
             "=Mitigation_Industry!R%dC" % (b + O_REV), h0 + TOT)
        line(r + 2, "tot", "", "Total rebate cost", "USD million", "obr", None, None, "", "=R%dC" % out[("obr", s)],
             h0 + TOT)
        line(r + 3, "tot", "", "Net revenue = total revenues - rebate cost", "USD million", "revnet", None,
             "revnet.cbam.tot.%d" % s, "", "=R[-2]C-R[-1]C", h0 + TOT)
        out[("rev", s)], out[("revnet", s)], out[("last", s)], out[("prod", s)] = r + 1, r + 3, r + 3, prod
        r += 6
    ws.Activate()
    wb.Application.ActiveWindow.FreezePanes = False
    ws.Range("L5").Select()
    wb.Application.ActiveWindow.FreezePanes = True
    return out


# ----------------------------------------------------------------------------------------------- kernel edit
def rewire_ener(wb, out):
    mi = wb.Worksheets("Mitigation_Industry")
    n = 0
    for s, hdrs in HEADERS.items():
        for h, sec in zip(hdrs, SECTORS):
            for i, f in enumerate(FUELS):
                r, a = h + ENER0 + i, out[("atpn", s, sec, f)]
                new = ener_template(h).replace(USAGE_OLD, USAGE_NEW.format(n=a))
                for c in DATA_COLS[1:]:
                    if mi.Cells(r, c).FormulaR1C1 != ener_template(h):
                        raise ValueError("ener formula changed at %s%d" % (col(c), r))
                    mi.Cells(r, c).FormulaR1C1 = new
                    n += 1
                old = mi.Cells(r, 9).Value
                note = "usage term reads net price Rebate_Industry row %d (atpn, Task I); efficiency term shp" % a
                mi.Cells(r, 9).Value = ("%s; %s" % (old, note)) if old else note
    return n


# ----------------------------------------------------------------------------------------------- Scenarios
def update_scenarios(wb, out):
    ws = wb.Worksheets("Scenarios")
    for r in (46, 47):
        if ws.Cells(r, 2).Value not in ("obrsh", "obrrb"):
            raise ValueError("Scenarios section 4 rows moved")
        ws.Cells(r, 8).Value = str(ws.Cells(r, 8).Value) + "; Stream 2 (I: Rebate_Industry)"
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task I (stream 2, %s): output-based rebate in sheet Rebate_Industry (scenario 1 rows %d-%d, scenario 2 rows "
        "%d-%d). Rebate per t = MIN(obrrb x EU benchmark, fuel + process carbon cost per t); sector rebate share rsh = "
        "rebate / carbon-cost base; net fuel price atpn = atp - (atp - ptp) x rsh replaces atp in the kernel ener USAGE "
        "term only (efficiency term shp and block process ER keep the full price). Codes egy.mit.atpn.<sector>.<fuel>"
        ".e.<s>, obr/ccf/rsh.ind.<sector>.<s>, obrbm/ccpt/obrpu/obrpt/ccnet/ppin/obr.<prod>.<s>, obr.cbam.tot.<s>, "
        "revnet.cbam.tot.<s>. ppin = net product price increase share = TASK H HOOK (production exogenous until Task H). "
        "Only bundles with theta > 0 (3B) change." % (VER, out[("first", 1)], out[("last", 1)], out[("first", 2)],
                                                      out[("last", 2)]))


# ----------------------------------------------------------------------------------------------- Check
def update_check(wb, out, ref):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task I (Stream2 v2): output-based rebating (identities expected 0, counts expected 0)"
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 2
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:J%d" % (hdr, hdr)).Value = ("Item", "", "", "Scenario", "Series", "max |diff| / count", "", "",
                                              "Expected", "")
    for j in range(10, 35):
        ws.Cells(hdr, j + 1).Value = 2011 + j
    maxabs = lambda r: "=MAX(MAX(%s),-MIN(%s))" % (rng(r), rng(r))
    sumf = lambda r: "=SUM(%s)" % rng(r)
    S = {}                      # stored-table range, filled before the pending rows are written
    pending = []

    def row(label, s, kind, fml, expected=0, agg=maxabs):
        pending.append((label, s, kind, fml, expected, agg))

    def stored(var, s, sec, bundle=True):   # INDEX/MATCH into the stored v1 table (key ends with the active bundle)
        key = '"v1:%s:%d:%s:"' % (var, s, sec) + ("&Settings!$B$10" if bundle else "")
        return 'INDEX({L}$%d:{L}$%d,MATCH(%s,$A$%d:$A$%d,0))' % (S["a"], S["b"], key, S["a"], S["b"])

    for s in (1, 2):
        prod, sh = out[("prod", s)], out[("obrsh", s)]
        pt0, pt1 = prod[("obrpt", 0)], prod[("obrpt", NP - 1)]
        cc0, cc1 = prod[("ccpt", 0)], prod[("ccpt", NP - 1)]
        nc0, nc1 = prod[("ccnet", 0)], prod[("ccnet", NP - 1)]
        ob0, ob1 = prod[("obr", 0)], prod[("obr", NP - 1)]
        row("Rebate per t while obrsh = 0: sum |obrpt| x (obrsh = 0), all products", s, "diff",
            lambda L, s=s, sh=sh, pt0=pt0, pt1=pt1: "=SUMPRODUCT(ABS(%s%s%d:%s%d))*(%s%s%d=0)"
            % (RI, L, pt0, L, pt1, RI, L, sh))
        row("Net carbon cost per t < 0 or > full carbon cost (count, all products)", s, "count",
            lambda L, nc0=nc0, nc1=nc1, cc0=cc0, cc1=cc1:
            "=SUMPRODUCT((%s%s%d:%s%d<-1E-9)+(%s%s%d:%s%d>%s%s%d:%s%d+1E-9))"
            % (RI, L, nc0, L, nc1, RI, L, nc0, L, nc1, RI, L, cc0, L, cc1), agg=sumf)
        for h, sec in zip(HEADERS[s], SECTORS):
            a0, a1 = out[("atpn", s, sec, FUELS[0])], out[("atpn", s, sec, FUELS[-1])]
            row("Net price atpn > atp (full price) or atpn < ptp (pre-tax) (count of fuel-years, %s)" % sec, s, "count",
                lambda L, h=h, a0=a0, a1=a1:
                "=SUMPRODUCT((%s%s%d:%s%d>%s%s%d:%s%d+1E-9)+(%s%s%d:%s%d<%s%s%d:%s%d-1E-9))"
                % (RI, L, a0, L, a1, MI, L, h + ATP0, L, h + ATP0 + NF - 1,
                   RI, L, a0, L, a1, MI, L, h + PTP0, L, h + PTP0 + NF - 1), agg=sumf)
        rs = [out[("rsh", s, sec)] for sec in SECTORS]
        row("Sector rebate share rsh outside [0,1] (count, 4 sectors)", s, "count",
            lambda L, rs=rs: "=" + "+".join("(%s%s%d<0)+(%s%s%d>1)" % (RI, L, x, RI, L, x) for x in rs), agg=sumf)
        row("Total rebate cost - sum of product rebate rows", s, "diff",
            lambda L, s=s, ob0=ob0, ob1=ob1: "=%s%s%d-SUM(%s%s%d:%s%d)" % (RI, L, out[("obr", s)], RI, L, ob0, L, ob1))
        row("Total rebate cost - sum of sector rebate rows", s, "diff",
            lambda L, s=s: "=%s%s%d-(%s)" % (RI, L, out[("obr", s)],
                                             "+".join("%s%s%d" % (RI, L, out[("obr", s, sec)]) for sec in SECTORS)))
        row("Net revenue - (total revenues - total rebate cost)", s, "diff",
            lambda L, s=s: "=%s%s%d-(%s%s%d-%s%s%d)"
            % (RI, L, out[("revnet", s)], RI, L, out[("rev", s)], RI, L, out[("obr", s)]))
    for h, sec in zip(HEADERS[2], SECTORS):
        row("Efficiency-term price unchanged: sum of shp over fuels (scenario 2, %s) - v1, active bundle" % sec, 2, "diff",
            lambda L, h=h, sec=sec: ("=SUM(%s%s%d:%s%d)-" % (MI, L, h + SHP0, L, h + SHP0 + NF - 1)
                                     + stored("shp", 2, sec).format(L=L)))
    for h, sec in zip(HEADERS[1], SECTORS):
        row("ener total scenario 1 (%s) - v1 (baseline, bundle-independent)" % sec, 1, "diff",
            lambda L, h=h, sec=sec: "=%s%s%d-" % (MI, L, h + TOT) + stored("ener", 1, sec, False).format(L=L))
    for h, sec in zip(HEADERS[2], SECTORS):
        row("ener total scenario 2 (%s) - v1 for the active bundle when its theta = 0 (else 0)" % sec, 2, "diff",
            lambda L, h=h, sec=sec: "=IF(Scenarios!$I$27=0,%s%s%d-%s,0)"
            % (MI, L, h + TOT, stored("ener", 2, sec).format(L=L)))
    n_ident = len(pending)
    for h, sec in zip(HEADERS[2], SECTORS):
        row("ener total scenario 2 (%s) below v1 while theta > 0 (count of years; the rebate can only raise energy use)"
            % sec, 2, "count",
            lambda L, h=h, sec=sec: "=IF(Scenarios!$I$27=0,0,--(%s%s%d<%s-1E-9))"
            % (MI, L, h + TOT, stored("ener", 2, sec).format(L=L)), agg=sumf)
    n_count = len(pending)
    for h, sec in zip(HEADERS[2], SECTORS):
        row("INFO: ener total scenario 2 (%s) - v1 for the active bundle (rebate effect, ktoe)" % sec, 2, "info",
            lambda L, h=h, sec=sec: "=%s%s%d-" % (MI, L, h + TOT) + stored("ener", 2, sec).format(L=L), expected="")
    row("INFO: total rebate cost scenario 2 (USD million)", 2, "info",
        lambda L: "=%s%s%d" % (RI, L, out[("obr", 2)]), expected="")
    row("INFO: net revenue scenario 2 (USD million)", 2, "info",
        lambda L: "=%s%s%d" % (RI, L, out[("revnet", 2)]), expected="")
    for sec in SECTORS:
        row("INFO: sector rebate share rsh scenario 2 (%s)" % sec, 2, "info",
            lambda L, sec=sec: "=%s%s%d" % (RI, L, out[("rsh", 2, sec)]), expected="")
    # stored v1 table position (after pending rows + note)
    r_first = hdr + 1
    r_note = r_first + len(pending) + 1
    S["a"] = r_note + 3
    S["b"] = S["a"] + len(SECTORS) * (1 + 2 * len(BUNDLES)) - 1
    r = r_first
    for label, s, kind, fml, expected, agg in pending:
        ws.Range("A%d" % r).Value = label
        ws.Range("D%d:E%d" % (r, r)).Value = (s, kind)
        for c in DATA_COLS:
            ws.Cells(r, c).Formula = fml(col(c))
        ws.Range("F%d" % r).Formula = agg(r)
        ws.Range("I%d" % r).Value = expected
        r += 1
    ws.Range("A%d" % r_note).Value = (
        "Note: for every bundle with theta = 0 (all but 3B) all kernel values equal v1 (atpn = atp exactly); for 3B the "
        "rebate lowers the net price in the ener usage term so energy use is >= v1 (INFO rows). The efficiency-term "
        "price (shp) is unchanged for all bundles. Production is exogenous until Task H.")
    ws.Range("A%d" % (S["a"] - 1)).Value = (
        "Stored v1 values (CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v1, computed per bundle by looping "
        "Settings!B10 before the Task I changes); key = v1:<var>:<scenario>:<sector>:<bundle>")
    ws.Range("A%d" % (S["a"] - 1)).Font.Bold = True
    r = S["a"]
    for sec in SECTORS:
        ws.Range("A%d:E%d" % (r, r)).Value = ("v1:ener:1:%s:" % sec, "tot", sec, 1, "v1 ener ktoe")
        ws.Range(rng(r)).Value = ref[("ener", 1, sec)]
        r += 1
    for b in BUNDLES:
        for sec in SECTORS:
            ws.Range("A%d:E%d" % (r, r)).Value = ("v1:ener:2:%s:%s" % (sec, b), "tot", sec, 2, "v1 ener ktoe (%s)" % b)
            ws.Range(rng(r)).Value = ref[("ener", 2, sec, b)]
            r += 1
        for sec in SECTORS:
            ws.Range("A%d:E%d" % (r, r)).Value = ("v1:shp:2:%s:%s" % (sec, b), "sum", sec, 2, "v1 sum shp (%s)" % b)
            ws.Range(rng(r)).Value = ref[("shp", 2, sec, b)]
            r += 1
    if r - 1 != S["b"]:
        raise ValueError("stored table size mismatch")
    ws.Range("%s%d:%s%d" % (L0, S["a"], L1, S["b"])).Interior.Color = GREEN
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task I)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (r_first, r_first + n_count - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    ws.Range("F%d" % (r0 + 1)).Value = "identities rows %d-%d, counts rows %d-%d" % (
        r_first, r_first + n_ident - 1, r_first + n_ident, r_first + n_count - 1)
    return r0 + 1


# ----------------------------------------------------------------------------------------------- Settings
def update_settings(wb, chk_row, out, n_cells):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: Task I section (row %d)." % (VER, chk_row)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task I (Stream 2, from Stream2_v1): output-based rebating. NEW SHEET Rebate_Industry (after Emissions_Industry; "
        "Mitigation_Industry layout; scenario 1 rows %d-%d, scenario 2 rows %d-%d): per product rebate per t = MIN(obrrb "
        "x EU benchmark [block col Z, Manual inputs Z30:Z37], fuel + process price increase per t [block o54-o61 + "
        "o65-o72]) (cap keeps the net price >= pre-tax price), net carbon cost per t, net product price increase share "
        "ppin (TASK H HOOK: wire into output; production is exogenous until Task H, so the output/usage effect acts only "
        "through the kernel energy usage term), rebate cost obr = rebate x production / 1000 (USD million) + total "
        "obr.cbam.tot; per sector rebate share rsh = MIN(1, sum rebate / sum carbon cost) and 9 net fuel prices atpn = "
        "atp - (atp - ptp) x rsh; net revenue revnet.cbam.tot = rev.cbam.tot - obr.cbam.tot. ONE FORMULA EDIT in "
        "Mitigation_Industry: the ener usage term IF(atp(t-1)<=0,1,atp(t)/atp(t-1))^eU now reads the Rebate_Industry "
        "atpn row (%d cells: columns M..AE of rows h+49..h+57 for headers 5/65/125/185 and 404/464/524/584, each "
        "asserted against the v1 R1C1 pattern first; column I note added); the efficiency term still reads shp (full "
        "price) and the block process-ER rows keep the full process price. MECHANISM CHOICE: explicit net-price rows "
        "rather than the shp/ssc wedge, because shp/ssc are separable per-fuel x per-sector series reserved for the Task "
        "J fund shadow price and carrying the rebate through them would require rewriting the atp rows; atp, shp, "
        "Data_Prices, Manual inputs and the CBAM block are untouched (block col Z benchmarks and o31-o38 / o43-o50 / "
        "o54-o61 / o65-o72 / o155 are now read). obrrb = theta x cptraj (Task E; ETS not included). Only bundle 3B "
        "(theta = 1) changes; LEGACY and every theta = 0 bundle reproduce v1 exactly (Check Task I section stores the v1 "
        "values per bundle; the build loops Settings!B10 over all bundles and asserts the section is 0). Merge = add the "
        "sheet, apply the same ener usage-term R1C1 edit, this row, the Check section, the Scenarios note and 'Read by' "
        "text; Sheets list above not extended (keeps row 25 fixed)."
        % (out[("first", 1)], out[("last", 1)], out[("first", 2)], out[("last", 2)], n_cells))
    ws.Range("C%d" % (last + 1)).WrapText = False


# ----------------------------------------------------------------------------------------------- verification
def verify(wb, chk_row, ref, out):
    xl, st, ck, ri = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Check"), wb.Worksheets("Rebate_Industry")
    mi = wb.Worksheets("Mitigation_Industry")
    c2030 = DATA_COLS[8]
    for b in BUNDLES:
        st.Range("B10").Value = b
        xl.CalculateFull()
        m = ck.Cells(chk_row, 4).Value
        msg = "bundle %-6s theta=%s  Task I max|identity|/count = %s" % (b, ref["theta"][b], m)
        if ref["theta"][b]:
            msg += "  | 2030: rsh %s; rebate %.1f, rev %.1f, revnet %.1f USD mn; ener s2 vs v1 %s ktoe" % (
                ["%.3f" % ri.Cells(out[("rsh", 2, sec)], c2030).Value for sec in SECTORS],
                ri.Cells(out[("obr", 2)], c2030).Value, ri.Cells(out[("rev", 2)], c2030).Value,
                ri.Cells(out[("revnet", 2)], c2030).Value,
                ["%.1f/%.1f" % (mi.Cells(h + TOT, c2030).Value, ref[("ener", 2, sec, b)][8])
                 for h, sec in zip(HEADERS[2], SECTORS)])
        print(msg)
        if m is None or abs(m) > 1e-9:
            raise ValueError("Task I check failed for bundle %s: %s" % (b, m))
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()


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
        ref = collect_v1(wb)
        out = build_rebate(wb)
        n = rewire_ener(wb, out)
        update_scenarios(wb, out)
        chk = update_check(wb, out, ref)
        update_settings(wb, chk, out, n)
        xl.CalculateFull()
        verify(wb, chk, ref, out)
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
