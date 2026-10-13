"""Build CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v3.xlsx from ..._Stream2_v2 (Task J: use-of-funds model).

Stream 2 branch (v0.7 -> Stream2_v1 Task F -> Stream2_v2 Task I -> this). Merge-friendly: one NEW sheet
(Fund_Industry), appended Check / Settings-log / Scenarios rows, formulas written into the Task E interface rows that
were created for exactly this purpose (Scenarios fundsp row, Data_Prices shp.<fuel>.ind.2 / ssc.<sector>.2, CBAM block
o4 / o5), and ONE controlled edit of the CBAM block process price term (o118-o125 and o15/o16, both scenarios: the
fund shadow price o4 is added inside the EXP(-beta x price) term; every cell is asserted against the v0.7 R1C1 pattern
first). No rows inserted anywhere; no Manual-inputs changes.

Mechanism (EgyptTaskReference.md section 2: fund size = share of revenue; fund -> shadow price on the efficiency
term, process emissions through the process semi-elasticity; applied AFTER the carbon price, not added to it)
  Fund size          F(t) = phi x rev0(t), phi = Scenarios bundle 'Fund share' (Data_Prices egy.mit.fundsh.<s>),
                     rev0 = CBAM block revenue at the carbon price BEFORE the fund's own abatement response
                     (revf0 = sum_k fuelCO2_k x covered_k x pf_k; revp0 = sum_k processCO2_k x covered_k x
                     EXP(-beta_k P_k) x P_k, with pf / P the block fuel / process carbon prices incl. ETS). Using rev0
                     instead of o155 breaks the circularity fund -> abatement -> revenue -> fund; the ex-post
                     revenue shortfall phi x o155 - F is reported (INFO).
  Budget -> price    the fund buys abatement along the block semi-elasticity curves, sequentially after the carbon
                     price: abatement A(s) = sum_k wf_k (1 - e^{-bf_k s}) + sum_k wp_k (1 - e^{-bp_k s}), with
                     wf_k = block o85 (fuel CO2 at baseline production, beta from o107-o114 column E) and
                     wp_k = o96 x covered x e^{-bp_k P_k} (process CO2 remaining after the carbon price, beta from
                     o118-o125 column E). Outlay rule (GREEN input, Fund_Industry!F4): 'RATE' = the fund pays a
                     uniform s $/t for every t abated (outlay = s x A(s), reverse-auction clearing price);
                     'COST' = the fund pays the abatement cost (outlay = integral of the marginal cost =
                     sum_k w_k [(1 - e^{-b s})/b - s e^{-b s}]). s solves outlay(s) = F by an explicit (non-circular)
                     safeguarded Newton iteration (15 rows, start s0 = sqrt(F / sum w b)), capped at smax (GREEN,
                     F5); capped years leave part of the fund unspent (reported).
  Delivery (block)   s -> Scenarios fundsp row -> Data_Prices egy.mit.fundsp.2 -> block o4 'Total Shadow Price'.
                     Fuel ER rows o107-o114 and o13/o14 already read o4 (Task E). Process ER rows o118-o125 and
                     o15/o16 now read EXP(-beta x (P + o4)): the fund's process abatement is the response at P + s
                     minus the response at P (sequential, not additive to the carbon price). o5 'Revenues Fund' =
                     F (Fund_Industry) instead of fundsh x o155 (which would be circular once o4 feeds back).
  Delivery (kernel)  Data_Prices egy.mit.shp.<fuel>.ind.2 = s x pass-through factor_f (Scenarios E30:E38, price units
                     per $/tCO2; 0 for bio / ren) and egy.mit.ssc.<sector>.2 = Fund_Industry ssc_fund (GREEN, F6,
                     default 1), so the kernel shadow price shp = atp + s x cpf reaches the EFFICIENCY term only;
                     the usage term keeps atpn / atp (Task I). Scenario 1 (baseline) is left as in CPAT (shp rows 0).
  LEGACY / phi = 0   F = 0 -> s = 0 -> every cell bit-identical to Stream2_v2 (checked per bundle). Only 3C (phi = 1)
                     changes.
Mechanism choice: shp/ssc wedge (reserved for this since Task E / Task I) rather than a new kernel column, because the
kernel formula F(t) = ... x (s(t)/s(t-1))^eF with s = atp + shpw x ssc already separates the efficiency price from
the usage price; the block already had o4 in the fuel-ER rows, so adding it to the process-ER price term is the
minimal completion of the Task E design. Note for the merge (Stream 1 v0.9 ERmax form of o118-o125 / Manual inputs
T:W rows 53-60): re-apply '+o4' inside the price term of whatever process-ER formula wins; the fund budget here is
calibrated on the ad hoc block semi-elasticities (Manual inputs E40:E47), not on the kernel eF energy response (Task
G reconciliation).

Added
  Fund_Industry    New sheet (after Rebate_Industry), Mitigation_Industry layout (A-AI, years K..AI, data L..AE).
                   Parameters F4:F6 (outlay rule, smax, ssc_fund). Per scenario: policy rows (fundsh, cptraj, etstraj,
                   pptraj), 8-product blocks pf, pp (carbon prices $/t), wf, wp (abatement bases Mt), bf, bp
                   (semi-elasticities), fund rows (revf0, revp0, rev0, fund F, emrp0, sum w b, s0), Newton rows
                   (A_k, outlay_k, dOutlay_k, s_k), result rows (fundsp s, abatement total / fuel / process, outlay,
                   residual, unspent, capped flag, cost per t, ex-post phi x o155, shortfall), kernel rows shpf
                   (9 fuels, s x cpf). Codes egy.mit.<var>.cbam.tot.<s>, egy.mit.<var>.<prod>.<s>,
                   egy.mit.shpf.<fuel>.ind.<s>. Scenario 1 is memo only (baseline fundsh = 0; not wired).
  Check            Section 'Task J (Stream2 v3)': budget residual = 0; block fund ER = -A(s) (exact identity);
                   rev0 = o155 when phi = 0; o5 = F; kernel shp - atp - s x cpf = 0; fundsp chain (Scenarios, o4) =
                   s; v2 preservation (ener totals, sum atp, o135, obr, revnet, eco2 for phi = 0 bundles, stored per
                   bundle); counts: s outside [0, smax], ener / o135 above v2 while phi > 0; INFO rows (F, s,
                   abatement, unspent, capped, ex-post shortfall, ener / eco2 / emissions change vs v2). The Task I
                   rows 'efficiency-term price unchanged vs v1' and 'ener = v1 when theta = 0' are conditioned on
                   phi = 0 too (they were only valid without a fund).
  Settings         Title, C25 note, version-log row 'v0.7branch_Stream2_v3'.
  Scenarios        fundsp row 49 L..AE = Fund_Industry s (was GREEN 0); note line 'Task J (stream 2, ...)'; 'Read by'
                   of fundsh / fundsp extended.

Run with Excel installed (from this folder):  python build_stream2_v3.py
The script loops Settings!B10 over all bundles (before the changes, to store v2 values; after, to verify the Task I
and Task J Check sections are 0 for each) and saves with Settings!B10 = LEGACY.
"""
import datetime
import os
import shutil

import win32com.client as win32

from build_v0_4 import BLUE, DATA_COLS, FUELS, GREEN, REVIEW, col, copy_formats
from build_stream2_v2 import (ATP0, BLOCK, BUNDLES, HEADERS, NF, NP, PRODUCTS, PROD_NAME, PROD_SECTOR, PTP0, SECTORS,
                              SHP0, TOT, L0, L1, rng)

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v2.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v3.xlsx")
VER = "v0.7branch_Stream2_v3"
MI, FI, EI, RI, SC = "Mitigation_Industry!", "Fund_Industry!", "Emissions_Industry!", "Rebate_Industry!", "Scenarios!"

# CBAM block offsets (block row o(n) = header + n)
O_CP, O_ETS, O_SP, O_FUND, O_FLAGF, O_PP, O_FLAGP = 2, 3, 4, 5, 54, 63, 65
O_FCO2, O_PCO2, O_ERF, O_ERP, O_EMRF, O_EMRP, O_EMIS, O_REV = 85, 96, 107, 118, 115, 126, 135, 155
O_NP, O_NO = 15, 16
N_NEWTON = 15
CPF_ROW0 = 30                                       # Scenarios E30:E38 pass-through factors, FUELS order
P_MODE, P_SMAX, P_SSC = 4, 5, 6                     # Fund_Industry parameter rows (value in column F)

ERP_OLD = "=R[-22]C*R[-33]C10*(EXP(-(RC5)*(R[-53]C6*R{pp}C+R[-53]C7*R{ets}C))-1)"
ERP_NEW = "=R[-22]C*R[-33]C10*(EXP(-(RC5)*(R[-53]C6*R{pp}C+R[-53]C7*R{ets}C+R{sp}C))-1)"
NPNO_OLD = ("=SUMPRODUCT(R{p0}C:R{p1}C*(R{e0}C{k}:R{e1}C{k}>0)*R{e0}C{k}:R{e1}C{k}*(1+R{c}C{k}*(EXP(-R{b0}C5:R{b1}C5"
            "*(R{f0}C6:R{f1}C6*R{pp}C+R{f0}C7:R{f1}C7*R{ets}C))-1)))/1000")
NPNO_NEW = NPNO_OLD.replace("R{f0}C7:R{f1}C7*R{ets}C))", "R{f0}C7:R{f1}C7*R{ets}C+R{sp}C))")


def erp_formulas(b):
    d = dict(pp=b + O_PP, ets=b + O_ETS, sp=b + O_SP)
    return ERP_OLD.format(**d), ERP_NEW.format(**d)


def npno_formulas(b, k):
    d = dict(p0=b + 31, p1=b + 38, e0=b + 19, e1=b + 26, k=k, c=b + 28, b0=b + O_ERP, b1=b + O_ERP + NP - 1,
             f0=b + O_FLAGP, f1=b + O_FLAGP + NP - 1, pp=b + O_PP, ets=b + O_ETS, sp=b + O_SP)
    return NPNO_OLD.format(**d), NPNO_NEW.format(**d)


def find_row(ws, column, value, first=1, last=2000):
    vals = ws.Range(ws.Cells(first, column), ws.Cells(last, column)).Value
    for i, (v,) in enumerate(vals):
        if v == value:
            return first + i
    raise ValueError("%s: no row with %s in column %d" % (ws.Name, value, column))


def find_prefix(ws, column, prefix, first=1, last=2000):
    vals = ws.Range(ws.Cells(first, column), ws.Cells(last, column)).Value
    return [first + i for i, (v,) in enumerate(vals) if isinstance(v, str) and v.startswith(prefix)]


# ----------------------------------------------------------------------------------------------- preconditions
def check_layout(wb):
    if "Fund_Industry" in [ws.Name for ws in wb.Worksheets]:
        raise ValueError("sheet Fund_Industry already exists")
    mi = wb.Worksheets("Mitigation_Industry")
    for s, hdrs in HEADERS.items():
        for h, sec in zip(hdrs, SECTORS):
            if mi.Cells(h, 3).Value != sec or mi.Cells(h, 4).Value != "Fuel Type" or mi.Cells(h, 10).Value != s:
                raise ValueError("sector header moved: row %d expected %s scenario %d" % (h, sec, s))
            if mi.Cells(h + 33, 6).Value != "ssc" or mi.Cells(h + 33, 7).Value != "egy.mit.ssc.%s.%d" % (sec, s):
                raise ValueError("ssc row moved: row %d" % (h + 33))
            for i, f in enumerate(FUELS):
                r = h + SHP0 + i
                if (mi.Cells(r, 6).Value != "shp" or mi.Cells(r, 7).Value != "egy.mit.shp.%s.ind.%d" % (f, s)
                        or mi.Cells(r, 20).FormulaR1C1
                        != "=R[-10]C+INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))*R%dC" % (h + 33)):
                    raise ValueError("shp row %d does not match the v0.7 pattern" % r)
            if mi.Cells(h + TOT, 2).Value != "tot" or mi.Cells(h + TOT, 6).Value != "ener":
                raise ValueError("ener total row moved: row %d" % (h + TOT))
    for s, b in BLOCK.items():
        if mi.Cells(b, 10).Value != s:
            raise ValueError("block header moved, scenario %d" % s)
        for off, label, code in ((O_SP, "Total Shadow Price", "egy.mit.fundsp.%d" % s),
                                 (O_FUND, "Revenues Fund", "egy.mit.fundsh.%d" % s),
                                 (O_CP, None, "egy.mit.cptraj.%d" % s), (O_ETS, None, "egy.mit.etstraj.%d" % s),
                                 (O_PP, None, "egy.mit.pptraj.%d" % s)):
            if (label and mi.Cells(b + off, 4).Value != label) or mi.Cells(b + off, 7).Value != code:
                raise ValueError("block o%d moved, scenario %d" % (off, s))
        if mi.Cells(b + O_FUND, 20).FormulaR1C1 != "=INDEX(Data_Prices!C,MATCH(RC7,Data_Prices!C1,0))*R[150]C":
            raise ValueError("block o5 formula changed, scenario %d" % s)
        for off, code in ((O_EMRF, "emrf"), (O_EMRP, "emrp"), (O_EMIS, "emis"), (O_REV, "rev")):
            if mi.Cells(b + off, 8).Value != "egy.mit.%s.cbam.tot.%d" % (code, s):
                raise ValueError("block o%d (%s) moved, scenario %d" % (off, code, s))
        if (mi.Cells(b + O_NP, 8).Value != "egy.mit.emisnp.cbam.tot.%d" % s
                or mi.Cells(b + O_NO, 8).Value != "egy.mit.emisno.cbam.tot.%d" % s):
            raise ValueError("block o15/o16 moved, scenario %d" % s)
        if mi.Cells(b + O_FCO2 - 1, 6).Value != "EF Fuel CO2" or mi.Cells(b + O_ERF - 1, 5).Value != "Semi-Elasticity" \
                or mi.Cells(b + O_ERP - 1, 5).Value != "Semi-Elasticity":
            raise ValueError("block o84 / o106 / o117 headers moved, scenario %d" % s)
        old_erp, _ = erp_formulas(b)
        for i in range(NP):
            if mi.Cells(b + O_FCO2 + i, 20).FormulaR1C1 != "=R[-54]C*RC6/1000":
                raise ValueError("block o%d (fuel CO2) formula changed" % (O_FCO2 + i))
            if mi.Cells(b + O_ERF + i, 20).FormulaR1C1 != "=R[-22]C*(EXP(-(RC5)*(R[-%d]C))-1)" % (O_ERF - O_SP + i):
                raise ValueError("block o%d (fuel ER) formula changed" % (O_ERF + i))
            for c in DATA_COLS:
                if mi.Cells(b + O_ERP + i, c).FormulaR1C1 != old_erp:
                    raise ValueError("block o%d (process ER) %s does not match the v0.7 pattern" % (O_ERP + i, col(c)))
        for off, k in ((O_NP, 10), (O_NO, 11)):
            old, _ = npno_formulas(b, k)
            for c in DATA_COLS:
                if mi.Cells(b + off, c).FormulaR1C1 != old:
                    raise ValueError("block o%d %s does not match the v0.7 pattern" % (off, col(c)))
    dp = wb.Worksheets("Data_Prices")
    for code in ["egy.mit.shp.%s.ind.2" % f for f in FUELS] + ["egy.mit.ssc.%s.2" % s for s in SECTORS]:
        r = find_row(dp, 1, code, 1, 200)
        vals = dp.Range(rng(r)).Value[0]
        if any(v not in (0, None) for v in vals):
            raise ValueError("Data_Prices %s is not 0 in %s..%s" % (code, L0, L1))
    if dp.Cells(1, 11).Value != 2021 or dp.Cells(1, 12).Value != 2022:
        raise ValueError("Data_Prices year columns moved")
    sc = wb.Worksheets("Scenarios")
    if ([sc.Cells(r, 2).Value for r in range(17, 24)] != list(BUNDLES) or sc.Cells(16, 10).Value != "Fund share (phi)"
            or sc.Cells(26, 10).Value != "Fund share (phi)"):
        raise ValueError("Scenarios bundle table moved")
    if (sc.Cells(49, 2).Value != "fundsp" or sc.Cells(48, 2).Value != "fundsh"
            or sc.Cells(49, 7).Value != "egy.mit.fundsp.2" or sc.Cells(49, 10).Value != "EGY"):
        raise ValueError("Scenarios section 4 rows moved")
    if any(v not in (0, None) for v in sc.Range("K49:AI49").Value[0]):
        raise ValueError("Scenarios fundsp row 49 is not the 0 placeholder")
    for i, f in enumerate(FUELS):
        if sc.Cells(CPF_ROW0 + i, 2).Value != f:
            raise ValueError("Scenarios pass-through factor rows moved")
    st = wb.Worksheets("Settings")
    if st.Range("B10").Value != "LEGACY":
        raise ValueError("source must be saved with Settings!B10 = LEGACY")
    if st.Range("A25").Value != "Check" or st.Range("A1").Value != "CPAT industry kernel - Egypt prototype v0.7branch_Stream2_v2":
        raise ValueError("Settings layout changed")
    ck = wb.Worksheets("Check")
    if not find_prefix(ck, 1, "Task I (Stream2 v2)"):
        raise ValueError("Check: Task I section not found")
    for sheet, code in (("Emissions_Industry", "egy.mit.eco2.irn.tot.e.2"), ("Rebate_Industry", "egy.mit.obr.cbam.tot.2"),
                        ("Rebate_Industry", "egy.mit.revnet.cbam.tot.2")):
        find_row(wb.Worksheets(sheet), 8, code, 1, 400)


# ----------------------------------------------------------------------------------------------- v2 reference values
def collect_v2(wb):
    """Per bundle (scenario 2): ener totals and sum of atp per sector, block o135 / o155, obr, revnet, kernel eco2
    total; plus bundle-independent scenario-1 ener totals."""
    xl, st, mi = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Mitigation_Industry")
    ei, ri, sc = wb.Worksheets("Emissions_Industry"), wb.Worksheets("Rebate_Industry"), wb.Worksheets("Scenarios")
    r_obr = find_row(ri, 8, "egy.mit.obr.cbam.tot.2", 1, 400)
    r_rnet = find_row(ri, 8, "egy.mit.revnet.cbam.tot.2", 1, 400)
    r_eco2 = [find_row(ei, 8, "egy.mit.eco2.%s.tot.e.2" % sec, 1, 400) for sec in SECTORS]
    ref = {"phi": {}, "rows": {"obr": r_obr, "revnet": r_rnet, "eco2": r_eco2}}
    b2 = BLOCK[2]
    for bd in BUNDLES:
        st.Range("B10").Value = bd
        xl.CalculateFull()
        ref["phi"][bd] = sc.Range("J27").Value
        for h, sec in zip(HEADERS[2], SECTORS):
            ref[("ener", 2, sec, bd)] = list(mi.Range(rng(h + TOT)).Value[0])
            atp = mi.Range("%s%d:%s%d" % (L0, h + ATP0, L1, h + ATP0 + NF - 1)).Value
            ref[("atp", 2, sec, bd)] = [sum(row[c] for row in atp) for c in range(len(DATA_COLS))]
        ref[("emis", 2, bd)] = list(mi.Range(rng(b2 + O_EMIS)).Value[0])
        ref[("rev", 2, bd)] = list(mi.Range(rng(b2 + O_REV)).Value[0])
        ref[("obr", 2, bd)] = list(ri.Range(rng(r_obr)).Value[0])
        ref[("revnet", 2, bd)] = list(ri.Range(rng(r_rnet)).Value[0])
        eco2 = [ei.Range(rng(r)).Value[0] for r in r_eco2]
        ref[("eco2", 2, bd)] = [sum(row[c] for row in eco2) for c in range(len(DATA_COLS))]
    for h, sec in zip(HEADERS[1], SECTORS):
        ref[("ener", 1, sec)] = list(mi.Range(rng(h + TOT)).Value[0])
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return ref


# ----------------------------------------------------------------------------------------------- Fund_Industry
def build_fund(wb):
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
    ws.Range("B2").Value = ("Mitigation Module - Egypt: Industry kernel - use-of-funds model (Task J, Stream 2): fund F "
                            "= phi x carbon revenue (pre-fund); fund shadow price s solves outlay(s) = F along the CBAM "
                            "block semi-elasticity curves (sequential, after the carbon price); s -> block o4 (fuel and "
                            "process ER) and -> kernel shp (efficiency term only) via Data_Prices shp/ssc")
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

    # parameters (GREEN inputs in column F)
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

        # per-product blocks: carbon prices, abatement bases and semi-elasticities (block-basis)
        blocks = (("pf", "Fuel-CO2 carbon price per t (block o54 flags x cptraj / ETS)", "$/ton CO2 real",
                   "=%sR{a}C6*R{cp}C+%sR{a}C7*R{ets}C" % (MI, MI), O_FLAGF, "block o54-o61 col F/G x o2/o3"),
                  ("pp", "Process-CO2 carbon price per t (block o65 flags x pptraj / ETS)", "$/ton CO2 real",
                   "=%sR{a}C6*R{pp}C+%sR{a}C7*R{ets}C" % (MI, MI), O_FLAGP, "block o65-o72 col F/G x o63/o3"),
                  ("wf", "Fuel CO2 abatement base (block fuel CO2 at the carbon price; fuel ER base)", "MtCO2",
                   "=%sR{a}C" % MI, O_FCO2, "block o85-o92"),
                  ("bf", "Fuel semi-elasticity beta_f (block o107-o114 col E)", "1/($/t)", "=%sR{a}C5" % MI, O_ERF,
                   "Manual inputs E40:E47"),
                  ("bp", "Process semi-elasticity beta_p (block o118-o125 col E)", "1/($/t)", "=%sR{a}C5" % MI, O_ERP,
                   "Manual inputs E40:E47"),
                  ("wp", "Process CO2 remaining after the carbon price = covered process CO2 x EXP(-beta_p x pp) "
                         "(process ER base for the fund)", "MtCO2",
                   "=%sR{a}C*%sR{c}C10*EXP(-R{bp}C*R{ppi}C)" % (MI, MI), O_PCO2, "block o96-o103 x o85 col J x EXP"))
        rows = {}
        for var, item, unit, tmpl, off, note in blocks:
            header(r, var, s, h0)
            rows[var] = r + 1
            for i, p in enumerate(PRODUCTS):
                f = tmpl.format(a=b + off + i, c=b + O_FCO2 + i, cp=r_cp, ets=r_ets, pp=r_pp,
                                bp=rows.get("bp", 0) + i, ppi=rows.get("pp", 0) + i)
                line(r + 1 + i, p, PROD_SECTOR[i], "%s - %s" % (item.split(" (")[0], PROD_NAME[i]), unit, var, None,
                     "%s.%s.%d" % (var, p, s), note if i == 0 else None, f, h0 + PTP0,
                     "0.000000" if var in ("bf", "bp") else "0.0000")
            r += NP + 2
        PF, PP, WF, BF, BP, WP = (rows[k] for k in ("pf", "pp", "wf", "bf", "bp", "wp"))

        def prod(tmpl, **kw):
            d = dict(pf0=PF, pf1=PF + NP - 1, pp0=PP, pp1=PP + NP - 1, wf0=WF, wf1=WF + NP - 1, bf0=BF,
                     bf1=BF + NP - 1, bp0=BP, bp1=BP + NP - 1, wp0=WP, wp1=WP + NP - 1, c0=b + O_FCO2,
                     c1=b + O_FCO2 + NP - 1, p0=b + O_PCO2, p1=b + O_PCO2 + NP - 1)
            d.update(kw)
            return tmpl.format(**d)

        # fund size (pre-fund revenue base) and Newton start
        header(r, "fund", s, h0)
        r_revf0, r_revp0, r_rev0, r_F, r_emrp0, r_swb, r_s0 = range(r + 1, r + 8)
        line(r_revf0, "revf0", "cbam", "Fuel-CO2 carbon revenue before the fund response (= block o137 at o4 = 0)",
             "USD million", "revf0", None, "revf0.cbam.tot.%d" % s, "sum wf x covered share (o85 col I) x pf",
             prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*%sR{c0}C9:R{c1}C9*R{pf0}C:R{pf1}C)" % MI), h0 + PTP0, "0.00")
        line(r_revp0, "revp0", "cbam", "Process-CO2 carbon revenue before the fund response (= block o138 at o4 = 0)",
             "USD million", "revp0", None, "revp0.cbam.tot.%d" % s, "sum wp x pp",
             prod("=SUMPRODUCT(R{wp0}C:R{wp1}C*R{pp0}C:R{pp1}C)"), h0 + PTP0, "0.00")
        line(r_rev0, "rev0", "cbam", "Carbon revenue base of the fund (pre-fund; = block o155 when phi = 0)",
             "USD million", "rev0", None, "rev0.cbam.tot.%d" % s,
             "revf0 + revp0; breaks the fund -> abatement -> revenue -> fund circularity (ex-post shortfall below)",
             "=R[-2]C+R[-1]C", h0 + PTP0, "0.00")
        line(r_F, "fund", "cbam", "Abatement fund F = phi x rev0", "USD million", "fund", None,
             "fund.cbam.tot.%d" % s, "feeds CBAM block o5 'Revenues Fund' (row %d)" % (b + O_FUND),
             "=R%dC*R[-1]C" % r_sh, h0 + PTP0, "0.00")
        line(r_emrp0, "emrp0", "cbam", "Process ER from the carbon price alone (= block o126 at o4 = 0)", "MtCO2",
             "emrp0", None, "emrp0.cbam.tot.%d" % s, "sum wp - sum covered process CO2",
             prod("=SUM(R{wp0}C:R{wp1}C)-SUMPRODUCT(%sR{p0}C:R{p1}C*%sR{c0}C10:R{c1}C10)" % (MI, MI)), h0 + PTP0)
        line(r_swb, "swb", "cbam", "Marginal abatement at s = 0: sum w x beta (fuel + process)", "MtCO2 per $/t",
             "swb", None, None, "A'(0); start value s0 = SQRT(F / A'(0)) (x2 for the COST rule)",
             prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*R{bf0}C:R{bf1}C)+SUMPRODUCT(R{wp0}C:R{wp1}C*R{bp0}C:R{bp1}C)"),
             h0 + PTP0, "0.000000")
        line(r_s0, "s0", "cbam", "Newton start value", "$/ton CO2 real", "s0", None, None, None,
             "=IF(OR(R%dC<=0,R%dC<=0),0,MIN(%s,SQRT(R%dC/R%dC*IF(%s=\"COST\",2,1))))"
             % (r_F, r_swb, SMAX, r_F, r_swb, MODE), h0 + PTP0, "0.0000")
        r = r_s0 + 2

        # safeguarded Newton iteration on outlay(s) = F (explicit rows, no circular reference)
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

        # results
        header(r, "result", s, h0)
        (r_s, r_A, r_Af, r_Ap, r_out, r_res, r_uns, r_cap, r_cpt, r_exp, r_sht) = range(r + 1, r + 12)
        line(r_s, "fundsp", "cbam", "Fund shadow price s (solution)", "$/ton CO2 real", "fundsp", None,
             "fundsp.cbam.tot.%d" % s, "feeds Scenarios fundsp row -> Data_Prices egy.mit.fundsp.%d -> CBAM block o4 "
             "(row %d); -> kernel shp via shpf rows below" % (s, b + O_SP), "=R%dC" % s_prev, h0 + PTP0, "0.0000")
        line(r_A, "abat", "cbam", "Fund-financed abatement A(s), fuel + process", "MtCO2", "abat", None,
             "abat.cbam.tot.%d" % s, "equals -(block o115 + o126 - emrp0) exactly (Check)", "=R[1]C+R[2]C", h0 + PTP0,
             "0.000000")
        line(r_Af, "abatf", "cbam", "Fund-financed fuel-CO2 abatement", "MtCO2", "abatf", None,
             "abatf.cbam.tot.%d" % s, "equals -block o115", prod("=SUMPRODUCT(R{wf0}C:R{wf1}C*(1-EXP(-R{bf0}C:R{bf1}C*R{s}C)))",
                                                          s=r_s), h0 + PTP0, "0.000000")
        line(r_Ap, "abatp", "cbam", "Fund-financed process-CO2 abatement", "MtCO2", "abatp", None,
             "abatp.cbam.tot.%d" % s, "equals emrp0 - block o126", prod("=SUMPRODUCT(R{wp0}C:R{wp1}C*(1-EXP(-R{bp0}C:R{bp1}C*R{s}C)))",
                                                                  s=r_s), h0 + PTP0, "0.000000")
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

        # kernel delivery: s in fuel price units
        header(r, "kernel", s, h0)
        ws.Range("I%d" % r).Value = ("s x Scenarios pass-through factor (E%d:E%d, fuel price unit per $/tCO2; bio / ren 0)"
                                     " -> Data_Prices egy.mit.shp.<fuel>.ind.%d -> kernel shp = atp + shpf x ssc "
                                     "(efficiency term only)" % (CPF_ROW0, CPF_ROW0 + NF - 1, s))
        for i, f in enumerate(FUELS):
            line(r + 1 + i, f, "ind", "Fund shadow price in fuel price units - %s" % f,
                 mi.Cells(h0 + SHP0 + i, 5).Value, "shpf", None, "shpf.%s.ind.%d" % (f, s), None,
                 "=R%dC*%sR%dC5" % (r_s, SC, CPF_ROW0 + i), h0 + SHP0 + i)
        out[("shpf", s)] = r + 1
        out[("last", s)] = r + NF
        for k, v in (("s", r_s), ("F", r_F), ("A", r_A), ("Af", r_Af), ("Ap", r_Ap), ("out", r_out), ("res", r_res),
                     ("uns", r_uns), ("cap", r_cap), ("exp", r_exp), ("sht", r_sht), ("rev0", r_rev0),
                     ("emrp0", r_emrp0), ("swb", r_swb)):
            out[(k, s)] = v
        r += NF + 4
    ws.Range("K3:AI%d" % r).HorizontalAlignment = -4152
    return out


# ----------------------------------------------------------------------------------------------- wiring
def wire_data_prices(wb, out):
    """Data_Prices shp.<fuel>.ind.2 = Fund_Industry shpf rows; ssc.<sector>.2 = ssc_fund parameter."""
    dp = wb.Worksheets("Data_Prices")
    n = 0
    for i, f in enumerate(FUELS):
        r = find_row(dp, 1, "egy.mit.shp.%s.ind.2" % f, 1, 200)
        dp.Range(rng(r)).FormulaR1C1 = "=%sR%dC" % (FI, out[("shpf", 2)] + i)
        dp.Cells(r, 7).Value = ("Shadow price of non-price policies ($/unit): Task J fund shadow price s x pass-through "
                                "factor (Fund_Industry shpf row %d); 0 when the bundle's fund share phi = 0"
                                % (out[("shpf", 2)] + i))
        n += len(DATA_COLS)
    for sec in SECTORS:
        r = find_row(dp, 1, "egy.mit.ssc.%s.2" % sec, 1, 200)
        dp.Range(rng(r)).FormulaR1C1 = "=%sR%dC6" % (FI, P_SSC)
        dp.Cells(r, 7).Value = ("Share of the shadow price added to the efficiency-term price shp: Task J ssc_fund "
                                "(Fund_Industry F%d, default 1)" % P_SSC)
        n += len(DATA_COLS)
    return n


def wire_scenarios(wb, out):
    ws = wb.Worksheets("Scenarios")
    ws.Range("K49:AI49").Interior.ColorIndex = -4142
    ws.Range(rng(49)).FormulaR1C1 = "=%sR%dC" % (FI, out[("s", 2)])
    ws.Range("K49").Value = 0
    ws.Range("AF49:AI49").Value = 0
    ws.Range("C49").Value = ("Shadow price of fund-financed abatement s (acts on intensity only): solved in "
                             "Fund_Industry so that fund outlay(s) = phi x carbon revenue (Task J)")
    ws.Range("H49").Value = ("CBAM block o4 -> fuel ER rows o107-o114 and process ER rows o118-o125 (price term P + s); "
                             "kernel shp via Data_Prices shp.<fuel>.ind.2 (Fund_Industry shpf rows)")
    ws.Range("H48").Value = str(ws.Range("H48").Value) + "; Fund_Industry fund size F = phi x rev0 (Task J)"
    last = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    ws.Range("B%d" % (last + 1)).Value = (
        "Task J (stream 2, %s): use-of-funds model in sheet Fund_Industry (scenario 1 rows %d-%d memo only, scenario 2 "
        "rows %d-%d). Fund F = phi x pre-fund carbon revenue rev0 (block fuel + process revenue at o4 = 0); the fund "
        "buys abatement along the CBAM block semi-elasticity curves after the carbon price: shadow price s solves "
        "outlay(s) = F (outlay rule RATE: s x A(s); alternative COST: area under the marginal cost curve; Fund_Industry "
        "F4:F6 parameters, cap smax). s -> fundsp row 49 -> block o4 (fuel ER o107-o114 as before; process ER "
        "o118-o125 and o15/o16 now use EXP(-beta x (P + o4))); o5 = F. Kernel: Data_Prices shp.<fuel>.ind.2 = s x "
        "pass-through factor (rows 30-38), ssc.<sector>.2 = ssc_fund, so shp = atp + s x cpf reaches the EFFICIENCY "
        "term only (usage term keeps atpn / atp). Codes egy.mit.fund/rev0/fundsp/abat/abatf/abatp/outlay/unspent/cpt"
        ".cbam.tot.<s>, pf/pp/wf/wp/bf/bp.<prod>.<s>, shpf.<fuel>.ind.<s>. Only bundles with phi > 0 (3C) change."
        % (VER, out[("first", 1)], out[("last", 1)], out[("first", 2)], out[("last", 2)]))


def wire_block(wb, out):
    """o5 = F; process ER rows o118-o125 and o15/o16 add o4 inside the EXP price term (both scenarios)."""
    mi = wb.Worksheets("Mitigation_Industry")
    n = 0
    for s, b in BLOCK.items():
        mi.Range(rng(b + O_FUND)).FormulaR1C1 = "=%sR%dC" % (FI, out[("F", s)])
        mi.Cells(b + O_FUND, 9).Value = ("Task J: fund F = fundsh x rev0 (Fund_Industry row %d; pre-fund revenue base, "
                                         "not o155, to avoid the fund -> abatement -> revenue circularity)" % out[("F", s)])
        mi.Cells(b + O_SP, 9).Value = ("Fund-financed abatement shadow price s (Task J, Fund_Industry row %d via "
                                       "Scenarios / Data_Prices); fuel ER o107-o114 and process ER o118-o125 (price "
                                       "term P + o4)" % out[("s", s)])
        n += len(DATA_COLS)
        old, new = erp_formulas(b)
        for i in range(NP):
            for c in DATA_COLS:
                cell = mi.Cells(b + O_ERP + i, c)
                if cell.FormulaR1C1 != old:
                    raise ValueError("block o%d %s changed" % (O_ERP + i, col(c)))
                cell.FormulaR1C1 = new
                n += 1
        mi.Cells(b + O_ERP, 9).Value = ("Task J: EXP(-beta x (process price + o4 fund shadow price)) - sequential: the "
                                        "fund's abatement is the response at P + s minus the response at P")
        for off, k in ((O_NP, 10), (O_NO, 11)):
            old, new = npno_formulas(b, k)
            for c in DATA_COLS:
                cell = mi.Cells(b + off, c)
                if cell.FormulaR1C1 != old:
                    raise ValueError("block o%d %s changed" % (off, col(c)))
                cell.FormulaR1C1 = new
                n += 1
        mi.Cells(b + O_NP, 9).Value = str(mi.Cells(b + O_NP, 9).Value or "") + " Task J: price term includes o4."
    return n


def fix_task_i_checks(wb):
    """The Task I rows 'efficiency-term price unchanged vs v1' and 'ener = v1 when theta = 0' hold only when the active
    bundle has no fund (phi = 0): condition them on Scenarios!$J$27 = 0 too."""
    ws = wb.Worksheets("Check")
    n = 0
    for r in find_prefix(ws, 1, "Efficiency-term price unchanged", 1400, 1700):
        for c in DATA_COLS:
            f = ws.Cells(r, c).Formula
            if not f.startswith("=SUM("):
                raise ValueError("Check row %d: unexpected formula" % r)
            ws.Cells(r, c).Formula = "=IF(Scenarios!$J$27=0,%s,0)" % f[1:]
        ws.Cells(r, 1).Value = str(ws.Cells(r, 1).Value) + " and phi = 0 (Task J)"
        n += 1
    for r in find_prefix(ws, 1, "ener total scenario 2 (", 1400, 1700):
        if "theta = 0" not in str(ws.Cells(r, 1).Value):
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


# ----------------------------------------------------------------------------------------------- Check
def update_check(wb, out, ref):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "Task J (Stream2 v3): use-of-funds model (identities expected 0, counts expected 0)"
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
    S = {}
    pending = []
    PHI0 = "Scenarios!$J$27=0"
    SMAX, SSC = "%s$F$%d" % (FI, P_SMAX), "%s$F$%d" % (FI, P_SSC)

    def row(label, s, kind, fml, expected=0, agg=maxabs):
        pending.append((label, s, kind, fml, expected, agg))

    def stored(var, s, sec, bundle=True):
        key = '"v2:%s:%d:%s:"' % (var, s, sec) + ("&Settings!$B$10" if bundle else "")
        return 'INDEX({L}$%d:{L}$%d,MATCH(%s,$A$%d:$A$%d,0))' % (S["a"], S["b"], key, S["a"], S["b"])

    def fi(k, s):
        return "%s{L}%d" % (FI, out[(k, s)])

    def mi(off, s):
        return "%s{L}%d" % (MI, BLOCK[s] + off)

    row("Fund shadow price s, scenario 1 (baseline fundsh = 0; memo block)", 1, "diff", lambda L: ("=" + fi("s", 1)).format(L=L))
    for s in (1, 2):
        row("Budget residual outlay(s) - F when F > 0, s uncapped and abatement available (Newton convergence)", s,
            "diff", lambda L, s=s: ("=%s*(%s>0)*(1-%s)*(%s>0)" % (fi("res", s), fi("F", s), fi("cap", s),
                                                                fi("swb", s))).format(L=L))
        row("Block fund ER identity: o115 + o126 - emrp0 + A(s) (fuel + process ER from the fund = -A exactly)", s,
            "diff", lambda L, s=s: ("=%s+%s-%s+%s" % (mi(O_EMRF, s), mi(O_EMRP, s), fi("emrp0", s), fi("A", s))).format(L=L))
        row("Pre-fund revenue base rev0 - block o155 when phi = 0 (else 0)", s, "diff",
            lambda L, s=s: ("=IF(%s=0,%s-%s,0)" % (fi("fundsh", s), fi("rev0", s), mi(O_REV, s))).format(L=L))
        row("Block o5 'Revenues Fund' - F", s, "diff", lambda L, s=s: ("=%s-%s" % (mi(O_FUND, s), fi("F", s))).format(L=L))
    for h, sec in zip(HEADERS[2], SECTORS):
        row("Kernel shadow-price wedge: sum over fuels of (shp - atp - shpf x ssc_fund), scenario 2, %s" % sec, 2,
            "diff", lambda L, h=h: ("=SUM(%s{L}%d:{L}%d)-SUM(%s{L}%d:{L}%d)-SUM(%s{L}%d:{L}%d)*%s"
                                    % (MI, h + SHP0, h + SHP0 + NF - 1, MI, h + ATP0, h + ATP0 + NF - 1, FI,
                                       out[("shpf", 2)], out[("shpf", 2)] + NF - 1, SSC)).format(L=L))
    row("Scenarios fundsp row 49 - s (scenario 2)", 2, "diff", lambda L: ("=%s{L}49-%s" % (SC, fi("s", 2))).format(L=L))
    row("Block o4 'Total Shadow Price' - s (scenario 2)", 2, "diff",
        lambda L: ("=%s-%s" % (mi(O_SP, 2), fi("s", 2))).format(L=L))
    for h, sec in zip(HEADERS[1], SECTORS):
        row("ener total scenario 1 (%s) - v2 (baseline, bundle-independent)" % sec, 1, "diff",
            lambda L, h=h, sec=sec: ("=%s{L}%d-" % (MI, h + TOT) + stored("ener", 1, sec, False)).format(L=L))
    for h, sec in zip(HEADERS[2], SECTORS):
        row("ener total scenario 2 (%s) - v2 for the active bundle when its phi = 0 (else 0)" % sec, 2, "diff",
            lambda L, h=h, sec=sec: ("=IF(%s,%s{L}%d-%s,0)" % (PHI0, MI, h + TOT, stored("ener", 2, sec))).format(L=L))
    for h, sec in zip(HEADERS[2], SECTORS):
        row("Usage-term price unchanged: sum of atp over fuels (scenario 2, %s) - v2, all bundles" % sec, 2, "diff",
            lambda L, h=h, sec=sec: ("=SUM(%s{L}%d:{L}%d)-%s" % (MI, h + ATP0, h + ATP0 + NF - 1,
                                                                stored("atp", 2, sec))).format(L=L))
    row("Block o135 total emissions scenario 2 - v2 when phi = 0 (else 0)", 2, "diff",
        lambda L: ("=IF(%s,%s-%s,0)" % (PHI0, mi(O_EMIS, 2), stored("emis", 2, "tot"))).format(L=L))
    row("Rebate cost obr.cbam.tot.2 - v2, all bundles (rebate independent of the fund)", 2, "diff",
        lambda L: ("=%s{L}%d-%s" % (RI, ref["rows"]["obr"], stored("obr", 2, "tot"))).format(L=L))
    row("Net revenue revnet.cbam.tot.2 - v2 when phi = 0 (else 0)", 2, "diff",
        lambda L: ("=IF(%s,%s{L}%d-%s,0)" % (PHI0, RI, ref["rows"]["revnet"], stored("revnet", 2, "tot"))).format(L=L))
    n_ident = len(pending)
    row("s outside [0, smax] (count of years, scenario 2)", 2, "count",
        lambda L: ("=(%s<-1E-9)+(%s>%s+1E-9)" % (fi("s", 2), fi("s", 2), SMAX)).format(L=L), agg=sumf)
    for h, sec in zip(HEADERS[2], SECTORS):
        row("ener total scenario 2 (%s) above v2 while phi > 0 (count of years; the fund can only lower energy use)"
            % sec, 2, "count",
            lambda L, h=h, sec=sec: ("=IF(%s,0,--(%s{L}%d>%s+1E-9))" % (PHI0, MI, h + TOT, stored("ener", 2, sec))).format(L=L),
            agg=sumf)
    row("Block o135 scenario 2 above v2 while phi > 0 (count of years)", 2, "count",
        lambda L: ("=IF(%s,0,--(%s>%s+1E-9))" % (PHI0, mi(O_EMIS, 2), stored("emis", 2, "tot"))).format(L=L), agg=sumf)
    n_count = len(pending)
    for k, label in (("F", "fund F (USD million)"), ("s", "fund shadow price s ($/tCO2)"),
                     ("A", "fund-financed abatement A (MtCO2)"), ("Af", "fuel-CO2 abatement (MtCO2)"),
                     ("Ap", "process-CO2 abatement (MtCO2)"), ("out", "fund outlay (USD million)"),
                     ("uns", "unspent fund (USD million)"), ("cap", "s capped at smax (1/0)"),
                     ("sht", "ex-post revenue shortfall phi x o155 - F (USD million)")):
        row("INFO: %s, scenario 2" % label, 2, "info", lambda L, k=k: ("=" + fi(k, 2)).format(L=L), expected="")
    for h, sec in zip(HEADERS[2], SECTORS):
        row("INFO: ener total scenario 2 (%s) - v2 for the active bundle (fund effect, ktoe)" % sec, 2, "info",
            lambda L, h=h, sec=sec: ("=%s{L}%d-" % (MI, h + TOT) + stored("ener", 2, sec)).format(L=L), expected="")
    row("INFO: kernel energy CO2 scenario 2 (4 sectors, Emissions_Industry) - v2 (MtCO2)", 2, "info",
        lambda L: ("=" + "+".join("%s{L}%d" % (EI, r) for r in ref["rows"]["eco2"]) + "-"
                   + stored("eco2", 2, "tot")).format(L=L), expected="")
    row("INFO: block o135 total emissions scenario 2 - v2 (MtCO2e)", 2, "info",
        lambda L: ("=%s-%s" % (mi(O_EMIS, 2), stored("emis", 2, "tot"))).format(L=L), expected="")
    row("INFO: block o155 total revenues scenario 2 - v2 (USD million)", 2, "info",
        lambda L: ("=%s-%s" % (mi(O_REV, 2), stored("rev", 2, "tot"))).format(L=L), expected="")
    row("INFO: net revenue revnet.cbam.tot.2 (USD million)", 2, "info",
        lambda L: ("=%s{L}%d" % (RI, ref["rows"]["revnet"])).format(L=L), expected="")
    # stored v2 table
    r_first = hdr + 1
    r_note = r_first + len(pending) + 1
    S["a"] = r_note + 3
    S["b"] = S["a"] + len(SECTORS) + len(BUNDLES) * (2 * len(SECTORS) + 6) - 1
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
        "Note: for every bundle with phi = 0 (all but 3C) all kernel and block values equal v2 (s = 0 exactly); for 3C "
        "the fund shadow price raises the efficiency-term price shp and the block ER rows, so energy use and emissions "
        "are <= v2 (INFO rows). atp (usage term) and the rebate are unchanged for all bundles. Residual tolerance: "
        "Newton converges to ~1E-12 USD million.")
    ws.Range("A%d" % (S["a"] - 1)).Value = (
        "Stored v2 values (CPAT_Industry_Kernel_Egypt_v0.7branch_Stream2_v2, computed per bundle by looping "
        "Settings!B10 before the Task J changes); key = v2:<var>:<scenario>:<sector|tot>:<bundle>")
    ws.Range("A%d" % (S["a"] - 1)).Font.Bold = True
    r = S["a"]
    for sec in SECTORS:
        ws.Range("A%d:E%d" % (r, r)).Value = ("v2:ener:1:%s:" % sec, "tot", sec, 1, "v2 ener ktoe")
        ws.Range(rng(r)).Value = ref[("ener", 1, sec)]
        r += 1
    for b in BUNDLES:
        for sec in SECTORS:
            ws.Range("A%d:E%d" % (r, r)).Value = ("v2:ener:2:%s:%s" % (sec, b), "tot", sec, 2, "v2 ener ktoe (%s)" % b)
            ws.Range(rng(r)).Value = ref[("ener", 2, sec, b)]
            r += 1
        for sec in SECTORS:
            ws.Range("A%d:E%d" % (r, r)).Value = ("v2:atp:2:%s:%s" % (sec, b), "sum", sec, 2, "v2 sum atp (%s)" % b)
            ws.Range(rng(r)).Value = ref[("atp", 2, sec, b)]
            r += 1
        for var, unit in (("emis", "MtCO2e"), ("rev", "USD million"), ("obr", "USD million"),
                          ("revnet", "USD million"), ("eco2", "MtCO2")):
            ws.Range("A%d:E%d" % (r, r)).Value = ("v2:%s:2:tot:%s" % (var, b), "tot", "cbam", 2, "v2 %s %s (%s)" % (var, unit, b))
            ws.Range(rng(r)).Value = ref[(var, 2, b)]
            r += 1
        ws.Range("A%d:E%d" % (r, r)).Value = ("v2:phi:2:tot:%s" % b, "phi", "cbam", 2, "v2 bundle fund share (info)")
        ws.Range(rng(r)).Value = ref["phi"][b]
        r += 1
    if r - 1 != S["b"]:
        raise ValueError("stored table size mismatch")
    ws.Range("%s%d:%s%d" % (L0, S["a"], L1, S["b"])).Interior.Color = GREEN
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (Task J)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=MAX(F%d:F%d)" % (r_first, r_first + n_count - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    ws.Range("F%d" % (r0 + 1)).Value = "identities rows %d-%d, counts rows %d-%d" % (
        r_first, r_first + n_ident - 1, r_first + n_ident, r_first + n_count - 1)
    return r0 + 1


# ----------------------------------------------------------------------------------------------- Settings
def update_settings(wb, chk_row, out, n_block, n_dp):
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: Task J section (row %d)." % (VER, chk_row)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Task J (Stream 2, from Stream2_v2): use-of-funds model. NEW SHEET Fund_Industry (after Rebate_Industry; "
        "Mitigation_Industry layout; parameters F4:F6 outlay rule RATE|COST, cap smax, ssc_fund; scenario 1 rows %d-%d "
        "MEMO ONLY (baseline fundsh = 0, not wired), scenario 2 rows %d-%d). Fund F = phi (Scenarios 'Fund share', "
        "egy.mit.fundsh.2) x rev0, rev0 = CBAM block fuel + process carbon revenue BEFORE the fund response (= o155 at "
        "o4 = 0; this breaks the fund -> abatement -> revenue -> fund circularity; the ex-post shortfall phi x o155 - F "
        "is reported). The fund buys abatement along the block semi-elasticity curves sequentially after the carbon "
        "price: A(s) = sum_k o85_k (1 - e^-bf s) + sum_k o96_k x covered x e^-bp P (1 - e^-bp s) with bf / bp = "
        "o107-o114 / o118-o125 col E (Manual inputs E40:E47); shadow price s solves outlay(s) = F by 15 explicit "
        "safeguarded Newton rows (RATE: outlay = s x A, uniform clearing price; COST: area under the marginal cost "
        "curve), capped at smax (unspent fund reported). DELIVERY: s -> Scenarios fundsp row 49 (was GREEN 0) -> "
        "Data_Prices egy.mit.fundsp.2 -> block o4 (fuel ER o107-o114 already read o4 since Task E). BLOCK EDITS (%d "
        "cells, both scenarios, each asserted against the v0.7 R1C1 first): process ER rows o118-o125 and o15/o16 now "
        "use EXP(-beta x (P + o4)) (sequential, not additive); o5 'Revenues Fund' = F from Fund_Industry instead of "
        "fundsh x o155 (would be circular). KERNEL: Data_Prices egy.mit.shp.<fuel>.ind.2 (rows 81-89) = s x Scenarios "
        "pass-through factor (E30:E38; bio / ren 0) and egy.mit.ssc.<sector>.2 = ssc_fund (%d cells; these CPAT-input "
        "rows were 0 placeholders reserved for this), so shp = atp + s x cpf reaches the kernel EFFICIENCY term only "
        "while the usage term keeps atpn / atp (Task I). MECHANISM CHOICE: the shp/ssc wedge (Task E design) rather "
        "than a new kernel column, because the kernel already separates the efficiency price (shp) from the usage price "
        "(atp); the block already had o4 in the fuel-ER rows, so adding it to the process-ER price term is the minimal "
        "completion. Check Task J section (row %d) stores the v2 values per bundle; the Task I rows 'efficiency-term "
        "price unchanged' and 'ener = v1 when theta = 0' are now also conditioned on phi = 0. Only bundle 3C (phi = 1) "
        "changes; LEGACY and every phi = 0 bundle reproduce v2 exactly. MERGE NOTES for Stream 1: (1) re-apply '+o4' "
        "inside the process price term of whatever o118-o125 / o15-o16 form wins (v0.9 ERmax form, Manual inputs T:W "
        "rows 53-60); (2) the fund budget is calibrated on the block semi-elasticities, not on the kernel eF energy "
        "response (Task G reconciliation); (3) o5 is redefined as the pre-fund fund size F; (4) Task H (Stream 1 v0.10) "
        "output response may change rev0 through production; (5) scenario-1 fund block is memo only. Sheets list above "
        "not extended (keeps row 25 fixed)." % (out[("first", 1)], out[("last", 1)], out[("first", 2)], out[("last", 2)],
                                               n_block, n_dp, chk_row))
    ws.Range("C%d" % (last + 1)).WrapText = False


# ----------------------------------------------------------------------------------------------- verification
def verify(wb, chk_row, ref, out):
    xl, st, ck = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Check")
    fi, mi = wb.Worksheets("Fund_Industry"), wb.Worksheets("Mitigation_Industry")
    ei = wb.Worksheets("Emissions_Industry")
    r_i = [r for r in find_prefix(ck, 1, "Max |identity|", 1, 2000) if "(Task I)" in ck.Cells(r, 1).Value]
    r_f = [r for r in find_prefix(ck, 1, "Max |identity|", 1, 2000) if "(Task F" in ck.Cells(r, 1).Value]
    if len(r_i) != 1:
        raise ValueError("Task I Max cell not found")
    res = {}
    for b in BUNDLES:
        st.Range("B10").Value = b
        xl.CalculateFull()
        m, mi_ = ck.Cells(chk_row, 4).Value, ck.Cells(r_i[0], 4).Value
        mf = ck.Cells(r_f[0], 4).Value if r_f else None
        msg = "bundle %-6s phi=%s  Task J max = %s | Task I max = %s | Task F max = %s" % (b, ref["phi"][b], m, mi_, mf)
        print(msg)
        if ref["phi"][b]:
            for yr, c in ((2030, DATA_COLS[8]), (2041, DATA_COLS[-1])):
                v = {k: fi.Cells(out[(k, 2)], c).Value for k in ("F", "s", "A", "Af", "Ap", "out", "uns", "cap", "sht")}
                v["ener"] = ["%.1f/%.1f" % (mi.Cells(h + TOT, c).Value, ref[("ener", 2, sec, b)][DATA_COLS.index(c)])
                             for h, sec in zip(HEADERS[2], SECTORS)]
                v["eco2"] = sum(ei.Cells(r, c).Value for r in ref["rows"]["eco2"]) - ref[("eco2", 2, b)][DATA_COLS.index(c)]
                v["o135"] = "%.3f/%.3f" % (mi.Cells(BLOCK[2] + O_EMIS, c).Value, ref[("emis", 2, b)][DATA_COLS.index(c)])
                v["o155"] = "%.1f/%.1f" % (mi.Cells(BLOCK[2] + O_REV, c).Value, ref[("rev", 2, b)][DATA_COLS.index(c)])
                res[(b, yr)] = v
                print("   %d: F %.1f, s %.2f, A %.3f (fuel %.3f, process %.3f), outlay %.1f, unspent %.2f, capped %s, "
                      "ex-post shortfall %.1f; ener s2 vs v2 %s; kernel eco2 delta %.3f; o135 %s; o155 %s"
                      % (yr, v["F"], v["s"], v["A"], v["Af"], v["Ap"], v["out"], v["uns"], v["cap"], v["sht"],
                         v["ener"], v["eco2"], v["o135"], v["o155"]))
        if m is None or abs(m) > 1e-9:
            raise ValueError("Task J check failed for bundle %s: %s" % (b, m))
        if mi_ is None or abs(mi_) > 1e-9:
            raise ValueError("Task I check failed for bundle %s: %s" % (b, mi_))
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    return res


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
        ref = collect_v2(wb)
        out = build_fund(wb)
        n_dp = wire_data_prices(wb, out)
        wire_scenarios(wb, out)
        n_block = wire_block(wb, out)
        fix_task_i_checks(wb)
        chk = update_check(wb, out, ref)
        update_settings(wb, chk, out, n_block, n_dp)
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
