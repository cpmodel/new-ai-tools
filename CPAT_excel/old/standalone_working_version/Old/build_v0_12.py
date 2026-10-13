"""Build CPAT_Industry_Kernel_Egypt_v0.12.xlsx from v0.11 (TODO T2: merge Stream 2 into the mainline).

Stream 1 = v0.8-v0.11 (Tasks K, D, H, L; CBAM block + Manual inputs). Stream 2 = branch off v0.7 with three increments
(Stream2_v1 Task F energy CO2, Stream2_v2 Task I output-based rebate, Stream2_v3 Task J use-of-funds). This builder
opens v0.11 and re-applies the three Stream-2 steps in order, reusing the branch builders' functions
(build_stream2_v1/v2/v3, kept importable from Old\\ after retirement), with these merge adaptations:

  Task F   applied unchanged (new sheets Data_EF, Emissions_Industry; MTInputs colouring; Check / Scenarios notes).
  Task I   applied unchanged (new sheet Rebate_Industry; kernel ener usage term reads atpn). Stored reference values
           are recomputed from v0.11 + Task F per bundle (not copied from the branch). Note: Task H (v0.10) already
           deducts obrrb from the cost that drives CBAM output (block o75-o82, basis = covered EF), so the Task I
           ppin 'Task H hook' stays MEMO (no double counting); Task I rebate cost (basis = EU benchmark) only feeds
           Rebate_Industry and the kernel usage term.
  Task J   process ER rows o118-o125 now have the v0.9 form ERmax x (1 - EXP(-beta x P)) per category (np: beta E,
           ERmax F; no: beta G, ERmax H; Manual inputs T:W rows 53-60). '+o4' is added inside both price terms of
           o118-o125 and inside the price term of o15/o16 (both scenarios; each cell asserted against v0.11 first).
           Fund_Industry process abatement is rebuilt per category: wm = Q x cov x MAX(EF,0) x ERmax / 1000
           (maximum abatable), wp = wm x EXP(-beta x P) (abatable remaining after the carbon price), bp = beta;
           16 rows per block (8 products x np, no). Then abatp(s) = sum wp (1 - e^{-bp s}) equals the change of
           block o126 exactly; revp0 = sum covered process CO2 x P - sum (wm - wp) x P (= o138 at o4 = 0);
           emrp0 = sum wp - sum wm (= o126 at o4 = 0). Fuel side unchanged (wf = o85, bf = o107 col E). Production
           (Task H) does not depend on o4 (o75-o82 use the o54/o65 cost rows), so no circularity.
  Settings one merged version-log row 'v0.12'; title; C25 note. Sheets list rows 17-25 not extended.

Verification (per bundle, Settings!B10 loop): every Check 'Max |...|' summary of Tasks A-L and F, I, J = 0; the
Stream-1 Check rows 2-4, 584-585, 1381 and every Mitigation_Industry value equal v0.11 for all bundles with
theta = phi = 0 (only 3B, theta = 1, and 3C, phi = 1, differ); no error values on any sheet.

Run with Excel installed (from this folder):  python build_v0_12.py
"""
import datetime
import os
import shutil
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))           # fallback if the Stream-2 builders are later moved to Old\

from build_v0_4 import BLUE, DATA_COLS, FUELS, GREEN, REVIEW, col, copy_formats   # noqa: E402
import build_stream2_v1 as s1                                                      # noqa: E402
import build_stream2_v2 as s2                                                      # noqa: E402
import build_stream2_v3 as s3                                                      # noqa: E402
from build_stream2_v2 import (BLOCK, BUNDLES, HEADERS, NF, NP, PRODUCTS, PROD_NAME, PROD_SECTOR, PTP0, SECTORS,  # noqa
                              SHP0, TOT, rng)

SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.11.xlsx")
if not os.path.exists(SRC):
    SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.11.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.12.xlsx")
VER = "v0.12"
s1.VER = s2.VER = s3.VER = VER
MI, FI = "Mitigation_Industry!", "Fund_Industry!"
O_CP, O_ETS, O_SP, O_FUND, O_FLAGF, O_PP, O_FLAGP = 2, 3, 4, 5, 54, 63, 65
O_COV, O_EF, O_PROD = 28, 19, 31
O_FCO2, O_PCO2, O_ERF, O_ERP, O_REV = 85, 96, 107, 118, 155
O_NP, O_NO = 15, 16
NPP = 2 * NP
P_MODE, P_SMAX, P_SSC = s3.P_MODE, s3.P_SMAX, s3.P_SSC
CPF_ROW0, N_NEWTON = s3.CPF_ROW0, s3.N_NEWTON
SNAP_SHEETS = ("Mitigation_Industry", "Check")


# ----------------------------------------------------------------------------------------------- snapshots
def snapshot(wb, last_check_row):
    """Values of Mitigation_Industry (whole used range) and the v0.11 Check rows, per bundle."""
    xl, st = wb.Application, wb.Worksheets("Settings")
    mi, ck = wb.Worksheets("Mitigation_Industry"), wb.Worksheets("Check")
    snap = {}
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        snap[("mi", bd)] = mi.Range("A1:AI%d" % mi.UsedRange.Rows.Count).Value
        snap[("ck", bd)] = ck.Range("A1:AI%d" % last_check_row).Value
        snap[("theta", bd)] = wb.Worksheets("Scenarios").Range("I27").Value
        snap[("phi", bd)] = wb.Worksheets("Scenarios").Range("J27").Value
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return snap


def maxdiff(a, b, skip=()):
    m, where = 0.0, None
    for i, (ra, rb) in enumerate(zip(a, b)):
        if i + 1 in skip:
            continue
        for j, (x, y) in enumerate(zip(ra, rb)):
            if isinstance(x, (int, float)) and isinstance(y, (int, float)):
                d = abs(x - y)
                if d > m:
                    m, where = d, (i + 1, j + 1)
            elif isinstance(x, (int, float)) != isinstance(y, (int, float)) and x not in (None, "") \
                    and y not in (None, ""):
                return float("inf"), (i + 1, j + 1, x, y)
    return m, where


# ----------------------------------------------------------------------------------------------- Task J preconditions
def fix_task_f_check(wb, chk_row):
    """Task F check (d) read block o1 (row 645), which is a header cascade to the year row (o29), not a price; in the
    branch this never bound because output was exogenous. Condition on the explicit price o2 + o3 (rows 646/647)."""
    ws = wb.Worksheets("Check")
    n = 0
    for r in range(chk_row, chk_row + 60):
        a = ws.Cells(r, 1).Value
        if not (a and str(a).startswith("Scenario 2 eco2 > scenario 1 eco2")):
            continue
        for c in DATA_COLS:
            L = col(c)
            f = ws.Cells(r, c).Formula
            old = "Mitigation_Industry!%s645>0" % L
            if f.count(old) != 1:
                raise ValueError("Check row %d %s: Task F (d) pattern not found" % (r, L))
            ws.Cells(r, c).Formula = f.replace(old, "(Mitigation_Industry!%s646+Mitigation_Industry!%s647)>0" % (L, L))
        ws.Cells(r, 1).Value = str(a).replace("block o1, row 645", "block o2 + o3, rows 646/647")
        n += 1
    if n != 4:
        raise ValueError("Task F (d) rows found: %d" % n)


def check_layout_j(wb):
    mi = wb.Worksheets("Mitigation_Industry")
    for s, hdrs in HEADERS.items():
        for h, sec in zip(hdrs, SECTORS):
            if mi.Cells(h + 33, 6).Value != "ssc" or mi.Cells(h + 33, 7).Value != "egy.mit.ssc.%s.%d" % (sec, s):
                raise ValueError("ssc row moved: row %d" % (h + 33))
            for i, f in enumerate(FUELS):
                r = h + SHP0 + i
                if (mi.Cells(r, 6).Value != "shp" or mi.Cells(r, 7).Value != "egy.mit.shp.%s.ind.%d" % (f, s)
                        or mi.Cells(r, 20).FormulaR1C1
                        != "=R[-10]C+INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))*R%dC" % (h + 33)):
                    raise ValueError("shp row %d does not match the v0.7 pattern" % r)
    for s, b in BLOCK.items():
        for off, code in ((O_SP, "fundsp"), (O_FUND, "fundsh"), (O_CP, "cptraj"), (O_ETS, "etstraj"), (O_PP, "pptraj")):
            if mi.Cells(b + off, 7).Value != "egy.mit.%s.%d" % (code, s):
                raise ValueError("block o%d moved, scenario %d" % (off, s))
        if mi.Cells(b + O_FUND, 20).FormulaR1C1 != "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))*R[150]C":
            raise ValueError("block o5 formula changed, scenario %d" % s)
        if mi.Cells(b + O_REV, 8).Value != "egy.mit.rev.cbam.tot.%d" % s:
            raise ValueError("block o155 moved, scenario %d" % s)
        if [mi.Cells(b + O_ERP - 1, c).Value for c in (5, 6, 7, 8)] != ["beta.np (1/$)", "ERmax.np", "beta.no (1/$)",
                                                                        "ERmax.no"]:
            raise ValueError("block o117 header (v0.9 ERmax form) changed, scenario %d" % s)
        if mi.Cells(b + O_FCO2 - 1, 6).Value != "EF Fuel CO2" or mi.Cells(b + O_ERF - 1, 5).Value != "Semi-Elasticity":
            raise ValueError("block o84 / o106 headers moved, scenario %d" % s)
        for i in range(NP):
            if mi.Cells(b + O_FCO2 + i, 20).FormulaR1C1 != "=R[-54]C*RC6/1000":
                raise ValueError("block o%d (fuel CO2) formula changed" % (O_FCO2 + i))
            if mi.Cells(b + O_ERF + i, 20).FormulaR1C1 != "=R[-22]C*(EXP(-(RC5)*(R[-%d]C))-1)" % (O_ERF - O_SP + i):
                raise ValueError("block o%d (fuel ER) formula changed" % (O_ERF + i))
            if mi.Cells(b + O_ERP + i, 20).FormulaR1C1 != erp_v011(b, i):
                raise ValueError("block o%d (process ER) does not match the v0.11 pattern:\n%s\n%s"
                                 % (O_ERP + i, mi.Cells(b + O_ERP + i, 20).FormulaR1C1, erp_v011(b, i)))
    dp = wb.Worksheets("Data_Prices")
    for code in ["egy.mit.shp.%s.ind.2" % f for f in FUELS] + ["egy.mit.ssc.%s.2" % s for s in SECTORS]:
        r = s3.find_row(dp, 1, code, 1, 200)
        if any(v not in (0, None) for v in dp.Range(rng(r)).Value[0]):
            raise ValueError("Data_Prices %s is not 0" % code)
    sc = wb.Worksheets("Scenarios")
    if sc.Cells(49, 2).Value != "fundsp" or sc.Cells(48, 2).Value != "fundsh" or sc.Cells(26, 10).Value != "Fund share (phi)":
        raise ValueError("Scenarios rows moved")
    if any(v not in (0, None) for v in sc.Range("K49:AI49").Value[0]):
        raise ValueError("Scenarios fundsp row 49 is not the 0 placeholder")
    for i, f in enumerate(FUELS):
        if sc.Cells(CPF_ROW0 + i, 2).Value != f:
            raise ValueError("Scenarios pass-through factor rows moved")


def erp_v011(b, i):
    """v0.11 R1C1 of process ER row o118+i (any data column)."""
    price = "(R%dC6*R%dC+R%dC7*R%dC)" % (b + O_FLAGP + i, b + O_PP, b + O_FLAGP + i, b + O_ETS)
    return ("=-R%dC/1000*(R%dC10*MAX(R%dC10,0)*RC6*(1-EXP(-RC5*%s))+R%dC11*MAX(R%dC11,0)*RC8*(1-EXP(-RC7*%s)))"
            % (b + O_PROD + i, b + O_COV, b + O_EF + i, price,
               b + O_COV, b + O_EF + i, price))


# ----------------------------------------------------------------------------------------------- Fund_Industry (v0.9 form)
def build_fund(wb):
    """build_stream2_v3.build_fund adapted to the v0.9 per-category ERmax process ER form."""
    mi = wb.Worksheets("Mitigation_Industry")
    ws = wb.Worksheets.Add(After=wb.Worksheets("Rebate_Industry"))
    ws.Name = "Fund_Industry"
    ws.Cells.Font.Name = "Arial"
    mi.Rows("1:2").Copy(ws.Rows(1))
    wb.Application.CutCopyMode = False
    for c in range(1, 36):
        ws.Columns(c).ColumnWidth = mi.Columns(c).ColumnWidth
    ws.Columns(4).ColumnWidth = 44
    ws.Columns(9).ColumnWidth = 60
    ws.Range("B2").Value = ("Mitigation Module - Egypt: Industry kernel - use-of-funds model (Task J): fund F = phi x "
                            "carbon revenue (pre-fund); fund shadow price s solves outlay(s) = F along the CBAM block "
                            "abatement curves (fuel semi-elasticities; process ERmax x (1 - e^-beta P) per category np / "
                            "no, Task D), sequential after the carbon price; s -> block o4 (fuel and process ER) and -> "
                            "kernel shp (efficiency term only) via Data_Prices shp/ssc")
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

    ws.Range("B%d" % (P_MODE - 1)).Value = "Fund parameters (GREEN = manual input)"
    ws.Range("B%d" % (P_MODE - 1)).Font.Bold = True
    params = ((P_MODE, "Outlay rule: RATE = fund pays s $/t for all abatement (reverse-auction clearing price); "
                       "COST = fund pays the abatement cost (area under the marginal cost curve)", "text", "RATE"),
              (P_SMAX, "Cap on the fund shadow price s (binding cap -> part of the fund stays unspent)",
                       "$/ton CO2 real", 1000),
              (P_SSC, "ssc_fund: share of s added to the kernel efficiency-term shadow price shp (Data_Prices "
                      "egy.mit.ssc.<sector>.2); 1 = full s (kernel usage term never sees s)", "share", 1))
    for r, label, unit, val in params:
        ws.Range("D%d" % r).Value = label
        ws.Range("E%d" % r).Value = unit
        ws.Range("F%d" % r).Value = val
        ws.Range("F%d" % r).Interior.Color = GREEN
        ws.Range("F%d" % r).HorizontalAlignment = -4152
    ws.Range("I%d" % P_MODE).Value = "Newton iterations: %d (explicit rows, no circular reference)" % N_NEWTON
    MODE, SMAX = "R%dC6" % P_MODE, "R%dC6" % P_SMAX

    dp = "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))"
    r = P_SSC + 2
    for s in (1, 2):
        b = BLOCK[s]
        h0 = HEADERS[s][0]
        copy_formats(mi.Rows(3), ws.Rows(r))
        wb.Application.CutCopyMode = False
        ws.Range("B%d" % r).Value = ("Abatement fund - %s (scenario %d); CBAM block rows %d-%d, kernel rows %d-%d%s"
                                     % ("Baseline" if s == 1 else "Policy", s, b, b + O_REV, h0, HEADERS[s][-1] + TOT,
                                        " - MEMO ONLY (baseline fundsh = 0; not wired to the block or the kernel)"
                                        if s == 1 else ""))
        r += 2
        header(r, "policy", s, h0)
        out[("first", s)] = r
        r_sh, r_cp, r_ets, r_pp = r + 1, r + 2, r + 3, r + 4
        line(r_sh, "fundsh", "", "Fund share of carbon revenue phi (active bundle)", "share", "fundsh",
             "egy.mit.fundsh.%d" % s, None, "Data_Prices (Scenarios section 4)", dp, h0 + PTP0)
        line(r_cp, "cptraj", "", "Carbon price on energy CO2", "$/ton CO2 real", "cptraj", None, None,
             "CBAM block o2 (row %d)" % (b + O_CP), "=%sR%dC" % (MI, b + O_CP), h0 + PTP0)
        line(r_ets, "etstraj", "", "ETS price", "$/ton CO2 real", "etstraj", None, None,
             "CBAM block o3 (row %d)" % (b + O_ETS), "=%sR%dC" % (MI, b + O_ETS), h0 + PTP0)
        line(r_pp, "pptraj", "", "Carbon price on process CO2", "$/ton CO2 real", "pptraj", None, None,
             "CBAM block o63 (row %d)" % (b + O_PP), "=%sR%dC" % (MI, b + O_PP), h0 + PTP0)
        out[("fundsh", s)] = r_sh
        r = r_pp + 2

        # product blocks; process blocks have 16 rows (8 products x category np, no)
        items8 = [(p, PROD_SECTOR[i], PROD_NAME[i], i, None) for i, p in enumerate(PRODUCTS)]
        items16 = [("%s.%s" % (p, cat), PROD_SECTOR[i], "%s (%s)" % (PROD_NAME[i], cat), i, cat)
                   for cat in ("np", "no") for i, p in enumerate(PRODUCTS)]
        rows = {}

        def blk(var, item, unit, items, fml, note, fmt="0.0000"):
            nonlocal r
            header(r, var, s, h0)
            rows[var] = r + 1
            for j, (key, sec, name, i, cat) in enumerate(items):
                line(r + 1 + j, key, sec, "%s - %s" % (item.split(" (")[0], name), unit, var, None,
                     "%s.%s.%d" % (var, key, s), note if j == 0 else None, fml(i, cat, j), h0 + PTP0, fmt)
            r += len(items) + 2

        blk("pf", "Fuel-CO2 carbon price per t (block o54 flags x cptraj / ETS)", "$/ton CO2 real", items8,
            lambda i, c, j: "=%sR%dC6*R%dC+%sR%dC7*R%dC" % (MI, b + O_FLAGF + i, r_cp, MI, b + O_FLAGF + i, r_ets),
            "block o54-o61 col F/G x o2/o3")
        blk("pp", "Process-CO2 carbon price per t (block o65 flags x pptraj / ETS)", "$/ton CO2 real", items16,
            lambda i, c, j: "=%sR%dC6*R%dC+%sR%dC7*R%dC" % (MI, b + O_FLAGP + i, r_pp, MI, b + O_FLAGP + i, r_ets),
            "block o65-o72 col F/G x o63/o3 (same price for np and no)")
        blk("wf", "Fuel CO2 abatement base (block o85-o92; fuel ER base)", "MtCO2", items8,
            lambda i, c, j: "=%sR%dC" % (MI, b + O_FCO2 + i), "block o85-o92")
        blk("bf", "Fuel semi-elasticity beta_f (block o107-o114 col E)", "1/($/t)", items8,
            lambda i, c, j: "=%sR%dC5" % (MI, b + O_ERF + i), "Manual inputs E40:E47", "0.000000")
        blk("bp", "Process semi-elasticity beta_p (block o118-o125 col E np / col G no)", "1/($/t)", items16,
            lambda i, c, j: "=%sR%dC%d" % (MI, b + O_ERP + i, 5 if c == "np" else 7),
            "Task D (Manual inputs T / V rows 53-60)", "0.000000")
        blk("wm", "Maximum abatable process CO2 = production x coverage x MAX(EF,0) x ERmax", "MtCO2", items16,
            lambda i, c, j: "=%sR%dC*%sR%dC%d*MAX(%sR%dC%d,0)*%sR%dC%d/1000"
            % (MI, b + O_PROD + i, MI, b + O_COV, 10 if c == "np" else 11, MI, b + O_EF + i, 10 if c == "np" else 11,
               MI, b + O_ERP + i, 6 if c == "np" else 8),
            "block o31 x o28 col J/K x o19 col J/K x o118 col F/H (Task D)")
        blk("wp", "Process CO2 abatable after the carbon price = wm x EXP(-beta_p x pp) (process ER base for the "
                  "fund)", "MtCO2", items16,
            lambda i, c, j: "=R%dC*EXP(-R%dC*R%dC)" % (rows["wm"] + j, rows["bp"] + j, rows["pp"] + j),
            "fund increment = wm e^-bP (1 - e^-bs)")
        PF, PP, WF, BF, BP, WM, WP = (rows[k] for k in ("pf", "pp", "wf", "bf", "bp", "wm", "wp"))

        def prod(tmpl, **kw):
            d = dict(pf0=PF, pf1=PF + NP - 1, pp0=PP, pp1=PP + NPP - 1, pp8=PP + NP - 1, wf0=WF, wf1=WF + NP - 1,
                     bf0=BF, bf1=BF + NP - 1, bp0=BP, bp1=BP + NPP - 1, wp0=WP, wp1=WP + NPP - 1, wm0=WM,
                     wm1=WM + NPP - 1, c0=b + O_FCO2, c1=b + O_FCO2 + NP - 1, p0=b + O_PCO2, p1=b + O_PCO2 + NP - 1)
            d.update(kw)
            return tmpl.format(**d)

        header(r, "fund", s, h0)
        r_revf0, r_revp0, r_rev0, r_F, r_emrp0, r_swb, r_s0 = range(r + 1, r + 8)
        line(r_revf0, "revf0", "cbam", "Fuel-CO2 carbon revenue before the fund response (= block o137 at o4 = 0)",
             "USD million", "revf0", None, "revf0.cbam.tot.%d" % s, "sum wf x covered share (o85 col I) x pf",
             prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*%sR{c0}C9:R{c1}C9*R{pf0}C:R{pf1}C)" % MI), h0 + PTP0, "0.00")
        line(r_revp0, "revp0", "cbam", "Process-CO2 carbon revenue before the fund response (= block o138 at o4 = 0)",
             "USD million", "revp0", None, "revp0.cbam.tot.%d" % s,
             "covered process CO2 (o96 x o85 col J) x pp - carbon-price abatement (wm - wp) x pp",
             prod("=SUMPRODUCT(%sR{p0}C:R{p1}C*%sR{c0}C10:R{c1}C10*R{pp0}C:R{pp8}C)"
                  "-SUMPRODUCT((R{wm0}C:R{wm1}C-R{wp0}C:R{wp1}C)*R{pp0}C:R{pp1}C)" % (MI, MI)), h0 + PTP0, "0.00")
        line(r_rev0, "rev0", "cbam", "Carbon revenue base of the fund (pre-fund; = block o155 when phi = 0)",
             "USD million", "rev0", None, "rev0.cbam.tot.%d" % s,
             "revf0 + revp0; breaks the fund -> abatement -> revenue -> fund circularity (ex-post shortfall below)",
             "=R[-2]C+R[-1]C", h0 + PTP0, "0.00")
        line(r_F, "fund", "cbam", "Abatement fund F = phi x rev0", "USD million", "fund", None,
             "fund.cbam.tot.%d" % s, "feeds CBAM block o5 'Revenues Fund' (row %d)" % (b + O_FUND),
             "=R%dC*R[-1]C" % r_sh, h0 + PTP0, "0.00")
        line(r_emrp0, "emrp0", "cbam", "Process ER from the carbon price alone (= block o126 at o4 = 0)", "MtCO2",
             "emrp0", None, "emrp0.cbam.tot.%d" % s, "sum wp - sum wm",
             prod("=SUM(R{wp0}C:R{wp1}C)-SUM(R{wm0}C:R{wm1}C)"), h0 + PTP0)
        line(r_swb, "swb", "cbam", "Marginal abatement at s = 0: sum w x beta (fuel + process)", "MtCO2 per $/t",
             "swb", None, None, "A'(0); start value s0 = SQRT(F / A'(0)) (x2 for the COST rule)",
             prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*R{bf0}C:R{bf1}C)+SUMPRODUCT(R{wp0}C:R{wp1}C*R{bp0}C:R{bp1}C)"),
             h0 + PTP0, "0.000000")
        line(r_s0, "s0", "cbam", "Newton start value", "$/ton CO2 real", "s0", None, None, None,
             "=IF(OR(R%dC<=0,R%dC<=0),0,MIN(%s,SQRT(R%dC/R%dC*IF(%s=\"COST\",2,1))))"
             % (r_F, r_swb, SMAX, r_F, r_swb, MODE), h0 + PTP0, "0.0000")
        r = r_s0 + 2

        A_T = prod("SUMPRODUCT(R{wf0}C:R{wf1}C*(1-EXP(-R{bf0}C:R{bf1}C*R{{s}}C)))"
                   "+SUMPRODUCT(R{wp0}C:R{wp1}C*(1-EXP(-R{bp0}C:R{bp1}C*R{{s}}C)))")
        AD_T = prod("(SUMPRODUCT(R{wf0}C:R{wf1}C*R{bf0}C:R{bf1}C*EXP(-R{bf0}C:R{bf1}C*R{{s}}C))"
                    "+SUMPRODUCT(R{wp0}C:R{wp1}C*R{bp0}C:R{bp1}C*EXP(-R{bp0}C:R{bp1}C*R{{s}}C)))")
        COST_T = prod("(SUMPRODUCT(R{wf0}C:R{wf1}C*((1-EXP(-R{bf0}C:R{bf1}C*R{{s}}C))/(R{bf0}C:R{bf1}C+(R{bf0}C:R{bf1}C=0))"
                      "-R{{s}}C*EXP(-R{bf0}C:R{bf1}C*R{{s}}C)*(R{bf0}C:R{bf1}C<>0)))"
                      "+SUMPRODUCT(R{wp0}C:R{wp1}C*((1-EXP(-R{bp0}C:R{bp1}C*R{{s}}C))/(R{bp0}C:R{bp1}C+(R{bp0}C:R{bp1}C=0))"
                      "-R{{s}}C*EXP(-R{bp0}C:R{bp1}C*R{{s}}C)*(R{bp0}C:R{bp1}C<>0))))")

        def outlay_f(rs, ra):
            return "=IF(%s=\"COST\",%s,R%dC*R%dC)" % (MODE, COST_T.format(s=rs), rs, ra)

        def grad_f(rs, ra):
            return "=IF(%s=\"COST\",R%dC*%s,R%dC+R%dC*%s)" % (MODE, rs, AD_T.format(s=rs), ra, rs, AD_T.format(s=rs))

        header(r, "newton", s, h0)
        ws.Range("I%d" % r).Value = ("%d safeguarded Newton steps on outlay(s) - F = 0; step k evaluates A, outlay "
                                     "and d outlay/ds at s(k-1); s(k) = s(k-1) - (outlay - F) / gradient, halved if "
                                     "negative, capped at smax; s = 0 when F <= 0" % N_NEWTON)
        s_prev = r_s0
        r += 1
        for k in range(1, N_NEWTON + 1):
            ra, ro, rg, rs = r, r + 1, r + 2, r + 3
            line(ra, "A%d" % k, "cbam", "Abatement A(s%d)" % (k - 1), "MtCO2", "A", None, None, None,
                 "=" + A_T.format(s=s_prev), h0 + PTP0, "0.000000")
            line(ro, "out%d" % k, "cbam", "Outlay(s%d)" % (k - 1), "USD million", "out", None, None, None,
                 outlay_f(s_prev, ra), h0 + PTP0, "0.0000")
            line(rg, "gd%d" % k, "cbam", "d outlay / ds at s%d" % (k - 1), "USD million per $/t", "gd", None, None,
                 None, grad_f(s_prev, ra), h0 + PTP0, "0.000000")
            step = "R%dC-(R%dC-R%dC)/R%dC" % (s_prev, ro, r_F, rg)
            line(rs, "s%d" % k, "cbam", "Shadow price after step %d" % k, "$/ton CO2 real", "s", None, None, None,
                 "=IF(R%dC<=0,0,IF(R%dC<=0,R%dC,MIN(%s,IF(%s<=0,R%dC/2,%s))))"
                 % (r_F, rg, s_prev, SMAX, step, s_prev, step), h0 + PTP0, "0.0000")
            s_prev = rs
            r += 4
        r += 1

        header(r, "result", s, h0)
        (r_s, r_A, r_Af, r_Ap, r_out, r_res, r_uns, r_cap, r_cpt, r_exp, r_sht) = range(r + 1, r + 12)
        line(r_s, "fundsp", "cbam", "Fund shadow price s (solution)", "$/ton CO2 real", "fundsp", None,
             "fundsp.cbam.tot.%d" % s, "feeds Scenarios fundsp row -> Data_Prices egy.mit.fundsp.%d -> CBAM block o4 "
             "(row %d); -> kernel shp via shpf rows below" % (s, b + O_SP), "=R%dC" % s_prev, h0 + PTP0, "0.0000")
        line(r_A, "abat", "cbam", "Fund-financed abatement A(s), fuel + process", "MtCO2", "abat", None,
             "abat.cbam.tot.%d" % s, "equals -(block o115 + o126 - emrp0) exactly (Check)", "=R[1]C+R[2]C", h0 + PTP0,
             "0.000000")
        line(r_Af, "abatf", "cbam", "Fund-financed fuel-CO2 abatement", "MtCO2", "abatf", None,
             "abatf.cbam.tot.%d" % s, "equals -block o115",
             prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*(1-EXP(-R{bf0}C:R{bf1}C*R{s}C)))", s=r_s), h0 + PTP0, "0.000000")
        line(r_Ap, "abatp", "cbam", "Fund-financed process-CO2 abatement (np + no)", "MtCO2", "abatp", None,
             "abatp.cbam.tot.%d" % s, "equals emrp0 - block o126",
             prod("=SUMPRODUCT(R{wp0}C:R{wp1}C*(1-EXP(-R{bp0}C:R{bp1}C*R{s}C)))", s=r_s), h0 + PTP0, "0.000000")
        line(r_out, "outlay", "cbam", "Fund outlay at s (RATE: s x A; COST: abatement cost)", "USD million", "outlay",
             None, "outlay.cbam.tot.%d" % s, None, outlay_f(r_s, r_A), h0 + PTP0, "0.0000")
        line(r_res, "resid", "cbam", "Budget residual outlay - F (0 unless capped or F <= 0)", "USD million", "resid",
             None, None, None, "=R%dC-R%dC" % (r_out, r_F), h0 + PTP0, "0.000000")
        line(r_uns, "unspent", "cbam", "Unspent fund = MAX(0, F - outlay)", "USD million", "unspent", None,
             "unspent.cbam.tot.%d" % s, None, "=MAX(0,R%dC-R%dC)" % (r_F, r_out), h0 + PTP0, "0.0000")
        line(r_cap, "capped", "cbam", "s at the cap smax (1/0)", "flag", "capped", None, None, None,
             "=(R%dC>0)*(R%dC>=%s-0.000000001)" % (r_F, r_s, SMAX), h0 + PTP0, "0")
        line(r_cpt, "cpt", "cbam", "Average fund cost per t abated = outlay / A", "$/ton CO2 real", "cpt", None,
             "cpt.cbam.tot.%d" % s, None, "=IF(R%dC>0,R%dC/R%dC,0)" % (r_A, r_out, r_A), h0 + PTP0, "0.0000")
        line(r_exp, "expost", "cbam", "Ex-post fund = phi x block o155 (revenue after the fund response)",
             "USD million", "expost", None, None, "INFO: o155 falls as the fund abates priced emissions",
             "=R%dC*%sR%dC" % (r_sh, MI, b + O_REV), h0 + PTP0, "0.00")
        line(r_sht, "short", "cbam", "Ex-post revenue shortfall = phi x o155 - F (<= 0)", "USD million", "short", None,
             None, "INFO", "=R[-1]C-R%dC" % r_F, h0 + PTP0, "0.00")
        r = r_sht + 2

        header(r, "kernel", s, h0)
        ws.Range("I%d" % r).Value = ("s x Scenarios pass-through factor (E%d:E%d, fuel price unit per $/tCO2; bio / ren 0)"
                                     " -> Data_Prices egy.mit.shp.<fuel>.ind.%d -> kernel shp = atp + shpf x ssc "
                                     "(efficiency term only)" % (CPF_ROW0, CPF_ROW0 + NF - 1, s))
        for i, f in enumerate(FUELS):
            line(r + 1 + i, f, "ind", "Fund shadow price in fuel price units - %s" % f,
                 mi.Cells(h0 + SHP0 + i, 5).Value, "shpf", None, "shpf.%s.ind.%d" % (f, s), None,
                 "=R%dC*%sR%dC5" % (r_s, "Scenarios!", CPF_ROW0 + i), h0 + SHP0 + i)
        out[("shpf", s)] = r + 1
        out[("last", s)] = r + NF
        for k, v in (("s", r_s), ("F", r_F), ("A", r_A), ("Af", r_Af), ("Ap", r_Ap), ("out", r_out), ("res", r_res),
                     ("uns", r_uns), ("cap", r_cap), ("exp", r_exp), ("sht", r_sht), ("rev0", r_rev0),
                     ("emrp0", r_emrp0), ("swb", r_swb)):
            out[(k, s)] = v
        r += NF + 4
    ws.Range("K3:AI%d" % r).HorizontalAlignment = -4152
    return out


def wire_block(wb, out):
    """o5 = F; '+o4' inside the process price term of o118-o125 (twice) and o15/o16 (once), both scenarios."""
    mi = wb.Worksheets("Mitigation_Industry")
    n = 0
    for s, b in BLOCK.items():
        mi.Range(rng(b + O_FUND)).FormulaR1C1 = "=%sR%dC" % (FI, out[("F", s)])
        mi.Cells(b + O_FUND, 9).Value = ("Task J: fund F = fundsh x rev0 (Fund_Industry row %d; pre-fund revenue base, "
                                         "not o155, to avoid the fund -> abatement -> revenue circularity)" % out[("F", s)])
        mi.Cells(b + O_SP, 9).Value = ("Fund-financed abatement shadow price s (Task J, Fund_Industry row %d via "
                                       "Scenarios / Data_Prices); fuel ER o107-o114, process ER o118-o125 and o15/o16 "
                                       "(price term P + o4)" % out[("s", s)])
        n += len(DATA_COLS)
        old_t, new_t = "C7*R%dC)" % (b + O_ETS), "C7*R%dC+R%dC)" % (b + O_ETS, b + O_SP)
        for row, cnt in [(b + O_ERP + i, 2) for i in range(NP)] + [(b + O_NP, 1), (b + O_NO, 1)]:
            ref = mi.Cells(row, DATA_COLS[0]).FormulaR1C1
            if ref.count(old_t) != cnt or "R%dC" % (b + O_SP) in ref:
                raise ValueError("block row %d: price term pattern not found %d times" % (row, cnt))
            if row >= b + O_ERP and ref != erp_v011(b, row - b - O_ERP):
                raise ValueError("block row %d changed" % row)
            for c in DATA_COLS:
                cell = mi.Cells(row, c)
                if cell.FormulaR1C1 != ref:
                    raise ValueError("block row %d %s differs from column %s" % (row, col(c), col(DATA_COLS[0])))
                cell.FormulaR1C1 = ref.replace(old_t, new_t)
                n += 1
        mi.Cells(b + O_ERP, 9).Value = ("Task J: price term = process price + o4 fund shadow price (np and no) - "
                                        "sequential: the fund's abatement is ERmax x e^-beta P x (1 - e^-beta s)")
        mi.Cells(b + O_NP, 9).Value = str(mi.Cells(b + O_NP, 9).Value or "") + " Task J: price term includes o4."
    return n


def fix_task_i_checks(wb):
    ws = wb.Worksheets("Check")
    n = 0
    for r in s3.find_prefix(ws, 1, "Efficiency-term price unchanged", 1, 6000):
        for c in DATA_COLS:
            f = ws.Cells(r, c).Formula
            if not f.startswith("=SUM("):
                raise ValueError("Check row %d: unexpected formula" % r)
            ws.Cells(r, c).Formula = "=IF(Scenarios!$J$27=0,%s,0)" % f[1:]
        ws.Cells(r, 1).Value = str(ws.Cells(r, 1).Value) + " and phi = 0 (Task J)"
        n += 1
    for r in s3.find_prefix(ws, 1, "ener total scenario 2 (", 1, 6000):
        if "v1 for the active bundle when its theta = 0" not in str(ws.Cells(r, 1).Value):
            continue
        for c in DATA_COLS:
            f = ws.Cells(r, c).Formula
            if not f.startswith("=IF(Scenarios!$I$27=0,"):
                raise ValueError("Check row %d: unexpected formula" % r)
            ws.Cells(r, c).Formula = f.replace("=IF(Scenarios!$I$27=0,", "=IF(AND(Scenarios!$I$27=0,Scenarios!$J$27=0),", 1)
        ws.Cells(r, 1).Value = str(ws.Cells(r, 1).Value) + " and phi = 0 (Task J)"
        n += 1
    if n != 8:
        raise ValueError("Task I check rows: expected 8, patched %d" % n)


def relabel_checks(wb):
    """Stored-reference labels: the merged build stores v0.11 (+ earlier merge steps) values, not branch files."""
    ws = wb.Worksheets("Check")
    for prefix, text in (("Stored v1 values", "Stored pre-Task-I values (v0.11 + Task F within the v0.12 build"),
                         ("Stored v2 values", "Stored pre-Task-J values (v0.11 + Tasks F, I within the v0.12 build")):
        rows = s3.find_prefix(ws, 1, prefix, 1, 6000)
        if len(rows) != 1:
            raise ValueError("Check: '%s' label not found" % prefix)
        v = ws.Cells(rows[0], 1).Value
        ws.Cells(rows[0], 1).Value = text + v[v.index(", computed"):]
    rows = s3.find_prefix(ws, 1, "Note: for every bundle with theta = 0", 1, 6000)
    if len(rows) != 1:
        raise ValueError("Check: Task I note not found")
    ws.Cells(rows[0], 1).Value = str(ws.Cells(rows[0], 1).Value).replace(
        "Production is exogenous until Task H.",
        "Production responds to the net carbon cost (Task H, v0.10, which already deducts obrrb); ppin is memo.")


def update_settings(wb, chk, out, n_ener, n_block, n_dp, diff_note):
    ws = wb.Worksheets("Settings")
    if ws.Range("A25").Value != "Check":
        raise ValueError("Settings layout changed")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + (
        " %s (Stream 2 merged): Task F section (row %d), Task I section (row %d), Task J section (row %d)."
        % (VER, chk["F"], chk["I"], chk["J"]))
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "T2 merge: Stream 2 (branch v0.7 -> Stream2_v1/v2/v3) integrated into the mainline v0.11 by re-applying the "
        "branch steps (build_v0_12.py imports build_stream2_v1/v2/v3). Task F: new sheets Data_EF, Emissions_Industry "
        "(eco2 = ener x EF), MTInputs EmissionsFactCO2 / EFsAdjustmentNonAnnexI green; check (d) now conditions on the "
        "explicit price o2 + o3 (branch used o1, a cascade to the year row). Task I: new sheet "
        "Rebate_Industry (Fund-independent output-based rebate per t = MIN(obrrb x EU benchmark, carbon cost per t), "
        "sector rebate share, net prices atpn); kernel ener usage term reads atpn (%d cells, asserted); ppin memo since "
        "Task H (v0.10) already deducts obrrb in the output response (basis covered EF; Task I basis EU benchmark -> "
        "CAVEATS). Task J: new sheet Fund_Industry (F = phi x pre-fund revenue rev0, shadow price s by 15 Newton rows, "
        "RATE|COST rule); process abatement rebuilt for the v0.9 per-category ERmax form (wm = Q x cov x EF x ERmax, wp "
        "= wm x e^-beta P, 16 rows np/no); block o5 = F, '+o4' inside the process price term of o118-o125 (np and no) "
        "and o15/o16 (%d block cells, asserted against v0.11); Data_Prices shp.<fuel>.ind.2 = s x pass-through, "
        "ssc.<sector>.2 = ssc_fund (%d cells); Scenarios fundsp row 49 = s. Check sections F (row %d), I (row %d), J "
        "(row %d) store the pre-step values per bundle computed inside this build. Regression vs v0.11: %s. Sheets list "
        "above not extended (keeps row 25 fixed)." % (n_ener, n_block, n_dp, chk["F"], chk["I"], chk["J"], diff_note))
    ws.Range("C%d" % (last + 1)).WrapText = False


# ----------------------------------------------------------------------------------------------- verification
def verify(wb, snap, last_check_row):
    xl, st, ck = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Check")
    mi = wb.Worksheets("Mitigation_Industry")
    max_rows = s3.find_prefix(ck, 1, "Max |", 1, 6000)
    errs, report = [], []
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        legacy_only = {2, 3, 4, 584, 585, 1381}
        bad = []
        for r in max_rows:
            v = ck.Cells(r, 4).Value
            label = ck.Cells(r, 1).Value
            if r in legacy_only:
                continue
            if not isinstance(v, (int, float)) or abs(v) > 1e-9:
                bad.append("%d %s = %s" % (r, label[:60], v))
        # ssc.<sector>.2 rows now carry ssc_fund (default 1); inert while shp.<fuel>.ind.2 = s x pass-through = 0
        ssc2 = {h + 33 for h in HEADERS[2]}
        d_mi, w_mi = maxdiff(snap[("mi", bd)], mi.Range("A1:AI%d" % len(snap[("mi", bd)])).Value, ssc2)
        d_ck, w_ck = maxdiff(snap[("ck", bd)], ck.Range("A1:AI%d" % last_check_row).Value)
        d_lo = max(abs((ck.Cells(r, 4).Value or 0) - (snap[("ck", bd)][r - 1][3] or 0)) for r in legacy_only)
        changed = bool(snap[("theta", bd)]) or bool(snap[("phi", bd)])
        msg = ("bundle %-6s theta=%s phi=%s | failing Max cells: %d | Mitigation_Industry max|v0.12-v0.11| = %.3g at %s"
               " | Check rows<=%d max diff = %.3g at %s | legacy-only summaries diff %.3g"
               % (bd, snap[("theta", bd)], snap[("phi", bd)], len(bad), d_mi, w_mi, last_check_row, d_ck, w_ck, d_lo))
        print(msg)
        report.append((bd, d_mi, d_ck))
        errs += ["%s: %s" % (bd, x) for x in bad]
        if not changed and (d_mi > 1e-9 or d_lo > 1e-9):
            errs.append("%s: differs from v0.11 (MI %.3g, legacy summaries %.3g)" % (bd, d_mi, d_lo))
        for ws in wb.Worksheets:
            ur = ws.UsedRange
            vals = ur.Value
            if not isinstance(vals, tuple):
                continue
            for i, rowv in enumerate(vals):
                for j, v in enumerate(rowv):
                    if isinstance(v, int) and v < -2146820000:
                        errs.append("%s: error value on %s!%s%d" % (bd, ws.Name, col(ur.Column + j), ur.Row + i))
                        break
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    if errs:
        raise ValueError("verification failed:\n" + "\n".join(errs[:40]))
    return report


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
        ck = wb.Worksheets("Check")
        last_ck = ck.Cells(ck.Rows.Count, 1).End(-4162).Row
        print("snapshot v0.11 ...")
        snap = snapshot(wb, last_ck)
        chk = {}
        # Task F
        print("Task F ...")
        s1.check_layout(wb)
        ef_rows = s1.build_data_ef(wb)
        rows = s1.build_emissions(wb)
        s1.colour_mtinputs(wb)
        s1.update_scenarios(wb, rows)
        chk["F"] = s1.update_check(wb, rows, ef_rows)
        fix_task_f_check(wb, chk["F"])
        # Task I
        print("Task I ...")
        s2.check_layout(wb)
        ref1 = s2.collect_v1(wb)
        out_i = s2.build_rebate(wb)
        n_ener = s2.rewire_ener(wb, out_i)
        s2.update_scenarios(wb, out_i)
        chk["I"] = s2.update_check(wb, out_i, ref1)
        xl.CalculateFull()
        s2.verify(wb, chk["I"], ref1, out_i)
        # Task J
        print("Task J ...")
        check_layout_j(wb)
        if not s3.find_prefix(ck, 1, "Task I (Stream2 v2)", 1, 6000):
            raise ValueError("Check: Task I section not found")
        ref2 = s3.collect_v2(wb)
        out_j = build_fund(wb)
        n_dp = s3.wire_data_prices(wb, out_j)
        s3.wire_scenarios(wb, out_j)
        n_block = wire_block(wb, out_j)
        fix_task_i_checks(wb)
        chk["J"] = s3.update_check(wb, out_j, ref2)
        relabel_checks(wb)
        xl.CalculateFull()
        s3.verify(wb, chk["J"], ref2, out_j)
        print("regression vs v0.11 ...")
        report = verify(wb, snap, last_ck)
        diff = "; ".join("%s %.3g" % (bd, d) for bd, d, _ in report if d > 1e-9) or "none"
        note = ("all Check summaries 0 for every bundle; Mitigation_Industry identical to v0.11 for LEGACY, 1A, 2A, 2B, "
                "3A (max |diff| 0); changed by design: %s" % diff)
        update_settings(wb, chk, out_j, n_ener, n_block, n_dp, note)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        if os.environ.get("KEEP_DEBUG"):
            try:
                wb.SaveAs(os.path.join(HERE, "_debug_v0_12.xlsx"))
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
