"""Build CPAT_Industry_Kernel_Egypt_v0.15.xlsx from v0.14 (TODO T5 Egypt CBAM EFs + TODO T1 Task D semi-elasticities).

T5  Manual inputs H30:K37 (S1 fc / fp / np / no) = Egypt CBAM EF v0.1 (egypt/supporting/EmissionFactors/
    EGY_CBAM_EF_v0.1.xlsx Products!D:G, CBAM conventions). Kernel conventions kept where they differ:
      urea  np = 0 (CBAM rule: urea-bound CO2 not deducted; was -0.733), M35 embedded NH3 = 1.1255904 (chain 1.2378);
      AN    no = 0.9944125 = integrated HNO3 N2O (0.79 t HNO3/t AN x 1.2588); M36 = 0.8462122 (all embedded NH3,
            direct + via HNO3); chain total 1.9528 (EF workbook stores the HNO3 N2O as a precursor instead).
    S:V memo (not read by any formula) restated as implied GJ/t = EF x 1000 / tCO2/TJ; AD/AE source prefix.
T1  Manual inputs E50 = IPCC; rows 53-60 Option A (ERmax 1, P* 100, ER* = IPCC AR6 central ER at $100, USD2019):
      DRI 0.3433, Scrap 0.1810, BF 0.2803, clinker 0.3352 (np); AN no 0.6179 (N2O-weighted blend of unabated 0.7210 /
      abated 0.3292 at the EF workbook's 50 % abatement); Al 0.2164 on np and no (combined central); ammonia and urea
      ERmax 0 (np = 0; fp routing of the CCS/blue lever not implemented). Q/R/S source text and confidence.
Fuel ER (rows 40-47, legacy half-elasticities on fc + fp) and the fund path are unchanged (spec step 5).
Task M stored 2030 snapshot (Table2_Industry section E) refilled; rebuild comparison column -> AdHoc rebuild v0.2.
Check: section 'T1/T5'; Settings version-log row; Scenarios note. Every bundle: all Check 'Max |' rows 0, no error
values; differences vs v0.14 are reported (expected), not asserted.

Run with Excel installed (from this folder):  python build_v0_15.py
"""
import datetime
import os

import win32com.client as win32

from build_stream2_v2 import ATP0, HEADERS, NF, SECTORS, TOT
from build_v0_4 import DATA_COLS, GREEN, REVIEW, col, copy_formats
from build_v0_14 import BUNDLES, SNAP, maxdiff, snapshot

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.14.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.15.xlsx")
VER = "v0.15"
MIS = "'Manual inputs'!"
REBUILD_K = {"1A": -31.79, "2A": -27.82, "2B": -29.34, "3A": -18.15, "3B": -12.74, "3C": -22.62}

# row: (fc, fp, np, no)  -- EGY_CBAM_EF_v0.1.xlsx Products!D:G (AN no: kernel convention, see docstring)
EF = {30: (0.17797725, 0.39301352, 0.03790102, 0.0),
      31: (0.04488, 0.0, 0.04398115, 0.0),
      32: (0.1683, 1.25668027, 0.05276542, 0.0),
      33: (0.3135685, 0.0, 0.53702616, 0.0),
      34: (0.71247, 1.26225, 0.0, 0.0),
      35: (0.1122, 0.0, 0.0, 0.0),
      36: (0.1122, 0.0, 0.0, 0.9944125),
      37: (0.12342, 0.0, 1.62349304, 0.774)}
EMB = {35: 1.1255904, 36: 0.8462122}
EF_SRC = ("v0.15 EF: Egypt CBAM EF v0.1 (egypt/supporting/EmissionFactors/EGY_CBAM_EF_v0.1.xlsx Products; "
          "EGY_CBAM_EF_Methodology_v0.1.md). ")
EF_NOTE = {35: "Urea np = 0 (CBAM rule: urea-bound CO2 not deducted; prototype -0.733). ",
           36: "AN no = integrated HNO3 N2O 0.79 x 1.2588 (kernel convention; EF workbook books it as precursor); "
               "50 % N2O abatement assumed (VERIFY). "}

SRC_TXT = ("IPCC AR6 WGIII Table 11.3/12.3 cost-bucket MACC, central (mean 2030 & long-run), USD2019; ER at $100 "
           "(Option A: ERmax 1, P* 100). Derivation: egypt/supporting/ProcessEmissions_CarbonPrice_Response; "
           "spec TASK-D v0.2. Lever: ")
# row: (np ERmax, np ER*, no ERmax, no ER*, np lever, no lever, confidence)
BETA = {53: (1, 0.3433, 0, 0, "H2/NG-DRI, scrap share, CCS on shaft", "-", "Medium"),
        54: (1, 0.1810, 0, 0, "electrode / carbon-input efficiency", "-", "Low"),
        55: (1, 0.2803, 0, 0, "top-gas recycling, DRI substitution, CCS", "-", "Medium"),
        56: (1, 0.3352, 0, 0, "clinker substitution (SCMs, LC3), CCS", "-", "Medium"),
        57: (0, 0, 0, 0, "none: np = 0 (SMR feedstock CO2 is fp; CCS / blue lever not routed to fp - caveat)", "-",
             "Low"),
        58: (0, 0, 0, 0, "none: inherits ammonia (np = 0 under CBAM rule)", "-", "Low"),
        59: (0, 0, 1, 0.6179, "-", "tertiary N2O abatement in HNO3 (blend: unabated 0.7210 / abated 0.3292 at "
             "50 % abated share, N2O-weighted)", "Medium"),
        60: (1, 0.2164, 1, 0.2164, "inert / improved anodes, PFC control (combined central, applied to np and no)",
             "inert / improved anodes, PFC control (combined central)", "Low")}


def write_inputs(wb):
    mi = wb.Worksheets("Manual inputs")
    for r, vals in EF.items():
        mi.Range("H%d:K%d" % (r, r)).Value = vals
        if r in EMB:
            mi.Range("M%d" % r).Value = EMB[r]
        t, v = mi.Range("T%d" % r).Value, mi.Range("V%d" % r).Value
        mi.Range("S%d" % r).Value = round(vals[0] * 1000 / t, 3) if t else 0
        mi.Range("U%d" % r).Value = round(vals[1] * 1000 / v, 3) if v else 0
        for c in ("AD", "AE"):
            old = str(mi.Range("%s%d" % (c, r)).Value or "")
            mi.Range("%s%d" % (c, r)).Value = EF_SRC + EF_NOTE.get(r, "") + "Prior: " + old
    if mi.Range("S27").Value in (None, ""):
        mi.Range("S27").Value = "S:V memo (v0.15): implied GJ/t = S1 EF x 1000 / tCO2 per TJ (not read by formulas)."
    mi.Range("E50").Value = "IPCC"
    for r, (e, g, i, kk, lnp, lno, conf) in BETA.items():
        mi.Range("E%d:G%d" % (r, r)).Value = (e, 100, g)
        mi.Range("I%d:K%d" % (r, r)).Value = (i, 100, kk)
        mi.Range("Q%d" % r).Value = SRC_TXT + lnp if e else "n/a (np ERmax 0): " + lnp
        mi.Range("R%d" % r).Value = SRC_TXT + lno if i else "n/a (no ERmax 0)"
        mi.Range("S%d" % r).Value = conf
        for c in ("E", "F", "G", "I", "J", "K"):
            mi.Range("%s%d" % (c, r)).Interior.Color = GREEN
        mi.Range("Q%d:S%d" % (r, r)).Interior.Color = GREEN


def fix_npno(wb):
    """EF-category rows o15 / o16 (Task B decomposition) applied the np beta to no and ignored ERmax; correct to the
    same form as the ER rows (np: ERmax F, beta E; no: ERmax H, beta G). Identical results while np and no had equal
    placeholder parameters (all prior versions)."""
    mi = wb.Worksheets("Mitigation_Industry")
    n = 0
    for b in (245, 644):
        e0, e1 = b + 118, b + 125
        for row, k, er, bt in ((b + 15, "J", "F", "E"), (b + 16, "K", "H", "G")):
            for c in DATA_COLS:
                f = mi.Cells(row, c).Formula
                old = "(1+$%s$%d*(EXP(-$E$%d:$E$%d*(" % (k, b + 28, e0, e1)
                new = "(1-$%s$%d*$%s$%d:$%s$%d*(1-EXP(-$%s$%d:$%s$%d*(" % (k, b + 28, er, e0, er, e1, bt, e0, bt, e1)
                if f.count(old) != 1 or not f.endswith("-1)))/1000"):
                    raise ValueError("unexpected formula Mitigation_Industry!%s%d: %s" % (col(c), row, f))
                mi.Cells(row, c).Formula = f.replace(old, new)[:-len("-1)))/1000")] + ")))/1000"
                n += 1
    return n


def refill_v2(wb):
    """Task J stored v2 table (Check, keys v2:<var>:<s>:<sec|tot>:<bundle>) = pre-fund values: recomputed per bundle
    with the bundle's fund share phi set to 0 (Scenarios J17:J23), same sources as build_stream2_v3.collect_v2."""
    xl, st, mi = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Mitigation_Industry")
    ei, ri, sc, ck = (wb.Worksheets(n) for n in ("Emissions_Industry", "Rebate_Industry", "Scenarios", "Check"))
    L0, L1 = col(DATA_COLS[0]), col(DATA_COLS[-1])
    rng = lambda r: "%s%d:%s%d" % (L0, r, L1, r)
    find = lambda ws, code: next(r for r in range(1, 401) if ws.Cells(r, 8).Value == code)
    r_obr, r_rnet = find(ri, "egy.mit.obr.cbam.tot.2"), find(ri, "egy.mit.revnet.cbam.tot.2")
    r_eco2 = [find(ei, "egy.mit.eco2.%s.tot.e.2" % sec) for sec in SECTORS]
    keys = {}
    for r in range(1, ck.UsedRange.Rows.Count + 1):
        a = ck.Cells(r, 1).Value
        if isinstance(a, str) and a.startswith("v2:"):
            keys[a] = r
    put = lambda key, vals: ck.Range(rng(keys[key])).__setattr__("Value", [list(vals)])
    tsum = lambda rows: [sum(x[c] for x in rows) for c in range(len(DATA_COLS))]
    n = 0
    for i, bd in enumerate(BUNDLES):
        phi = sc.Range("J%d" % (17 + i)).Value
        sc.Range("J%d" % (17 + i)).Value = 0
        st.Range("B10").Value = bd
        xl.CalculateFull()
        if sc.Range("J27").Value != 0:
            raise ValueError("phi not 0 for " + bd)
        for h, sec in zip(HEADERS[2], SECTORS):
            put("v2:ener:2:%s:%s" % (sec, bd), mi.Range(rng(h + TOT)).Value[0])
            put("v2:atp:2:%s:%s" % (sec, bd), tsum(mi.Range("%s%d:%s%d" % (L0, h + ATP0, L1, h + ATP0 + NF - 1)).Value))
            n += 2
        b2 = 644
        for var, src in (("emis", mi.Range(rng(b2 + 135)).Value[0]), ("rev", mi.Range(rng(b2 + 155)).Value[0]),
                         ("obr", ri.Range(rng(r_obr)).Value[0]), ("revnet", ri.Range(rng(r_rnet)).Value[0]),
                         ("eco2", tsum([ei.Range(rng(r)).Value[0] for r in r_eco2]))):
            put("v2:%s:2:tot:%s" % (var, bd), src)
            n += 1
        if bd == "LEGACY":
            for h, sec in zip(HEADERS[1], SECTORS):
                put("v2:ener:1:%s:" % sec, mi.Range(rng(h + TOT)).Value[0])
                n += 1
        sc.Range("J%d" % (17 + i)).Value = phi
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    note = [r for r in range(1, ck.UsedRange.Rows.Count + 1)
            if str(ck.Cells(r, 1).Value).startswith("Stored v2 values") or
            str(ck.Cells(r, 1).Value).startswith("Stored pre-Task-J values")]
    for r in note:
        ck.Cells(r, 1).Value = str(ck.Cells(r, 1).Value) + " [refilled by build_v0_15.py: phi = 0 per bundle, T1/T5 inputs]"
    return n


def t2_rows(wb):
    ws = wb.Worksheets("Table2_Industry")
    k = {"snap": []}
    for r in range(1, ws.UsedRange.Rows.Count + 1):
        a = ws.Range("A%d" % r).Value
        if a == "live":
            k["live"] = r
        elif a == "stored":
            k["snap"].append(r)
    hdr = k["live"] - 1
    k["kc"] = next(c for c in range(1, 40) if str(ws.Cells(hdr, c).Value).startswith("Rebuild v0.1 K"))
    k["n"] = k["kc"] - 6
    return k


def refill_table2(wb, k):
    xl, st, ws = wb.Application, wb.Worksheets("Settings"), wb.Worksheets("Table2_Industry")
    hdr = k["live"] - 1
    ws.Cells(hdr, k["kc"]).Value = "Rebuild v0.2 K"
    for r in range(k["snap"][-1] + 1, k["snap"][-1] + 4):
        v = ws.Range("B%d" % r).Value
        if isinstance(v, str) and "Rebuild v0.1" in v:
            ws.Range("B%d" % r).Value = (v.replace("build_v0_14.py", "build_v0_14.py, refilled by build_v0_15.py")
                                         .replace("Rebuild v0.1", "Rebuild v0.2").replace("v0.1 section 3", "v0.2"))
    last = col(5 + k["n"])
    for r in k["snap"]:
        bd = ws.Range("B%d" % r).Value
        ws.Cells(r, k["kc"]).Value = REBUILD_K[bd]
        st.Range("B10").Value = bd
        xl.CalculateFull()
        ws.Range("C%d:E%d" % (r, r)).Value = ws.Range("C%d:E%d" % (k["live"], k["live"])).Value
        ws.Range("F%d:%s%d" % (r, last, r)).Value = ws.Range("F%d:%s%d" % (k["live"], last, k["live"])).Value
    st.Range("B10").Value = "LEGACY"
    xl.CalculateFull()
    ck = wb.Worksheets("Check")
    for r in range(1, ck.UsedRange.Rows.Count + 1):
        v = ck.Range("A%d" % r).Value
        if isinstance(v, str) and "TASK-2b rebuild v0.1" in v:
            ck.Range("A%d" % r).Value = v.replace("rebuild v0.1", "rebuild v0.2")


def update_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "T1/T5 (%s): Egypt CBAM EF v0.1 and Task D semi-elasticities" % VER
    ws.Range("A%d" % r0).Font.Bold = True
    hdr = r0 + 3
    copy_formats(ws.Rows(5), ws.Rows(hdr))
    wb.Application.CutCopyMode = False
    ws.Range("A%d:I%d" % (hdr, hdr)).Value = ("Item", "", "", "", "Kind", "|diff| / count", "", "", "Expected")
    items = []
    for r, vals in EF.items():
        items.append(("H%d:K%d (%s) = EGY_CBAM_EF_v0.1 Products D:G" % (r, r, wb.Worksheets("Manual inputs")
                                                                          .Range("D%d" % r).Value),
                      "=" + "+".join("ABS(%s%s%d-%.10g)" % (MIS, c, r, v) for c, v in zip("HIJK", vals))))
    for r, v in EMB.items():
        items.append(("M%d embedded NH3 = %.7g" % (r, v), "=ABS(%sM%d-%.10g)" % (MIS, r, v)))
    items.append(("Task D applied: E50 = IPCC", '=IF(%sE50="IPCC",0,1)' % MIS))
    items.append(("Task D applied: ER(100) row 53 = 0.3433 +/- 1e-6",
                  "=IF(ABS(%sU53*(1-EXP(-%sT53*100))-0.3433)<=1E-6,0,1)" % (MIS, MIS)))
    for r, (e, g, i, kk, *_rest) in BETA.items():
        items.append(("Row %d: ER(P*) used = ER* (np %.4f, no %.4f)" % (r, g, kk),
                      "=ABS(%sU%d*(1-EXP(-%sT%d*100))-%.10g)+ABS(%sW%d*(1-EXP(-%sV%d*100))-%.10g)"
                      % (MIS, r, MIS, r, g, MIS, r, MIS, r, kk)))
    r = hdr + 1
    for label, f in items:
        ws.Range("A%d" % r).Value = label
        ws.Range("E%d" % r).Value = "diff"
        ws.Range("F%d" % r).Formula = f
        ws.Range("I%d" % r).Value = 0
        r += 1
    ws.Range("A%d" % (r0 + 1)).Value = "Max |identity| and violation counts (T1/T5)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=IF(MAX(F%d:F%d)<1E-9,0,MAX(F%d:F%d))" % (hdr + 1, r - 1, hdr + 1, r - 1)
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    return r0 + 1


def update_notes(wb, chk, stats):
    sc = wb.Worksheets("Scenarios")
    last = sc.Cells(sc.Rows.Count, 2).End(-4162).Row
    sc.Range("B%d" % (last + 1)).Value = (
        "T1/T5 (%s): Manual inputs H30:K37 = Egypt CBAM EF v0.1 (urea np 0; AN no = integrated HNO3 N2O 0.9944); "
        "E50 = IPCC with Task D rows 53-60 (IPCC AR6 central ER at $100; AN blend 0.6179; Al 0.2164 np+no; ammonia / "
        "urea 0). Table 2 2030 snapshot refilled." % VER)
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: T1/T5 section (row %d)." % (VER, chk)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "T5: S1 emission factors H30:K37 = Egypt CBAM EF v0.1 (EGY_CBAM_EF_v0.1.xlsx; urea np 0 CBAM rule, M35 "
        "1.1256; AN no 0.9944 integrated HNO3 N2O, M36 0.8462); S:V memo restated. T1: Task D semi-elasticities "
        "applied (E50 IPCC; rows 53-60 ER at $100 = IPCC AR6 central, USD2019: DRI 0.3433, scrap 0.1810, BF 0.2803, "
        "clinker 0.3352, AN no 0.6179 (50 %% abatement blend), Al 0.2164 np+no, ammonia/urea 0). Fuel ER (rows 40-47) "
        "and fund unchanged. Table 2 snapshot refilled; rebuild column = AdHoc rebuild v0.2. Check row %d. 2030 K: %s."
        % (chk, stats))
    ws.Range("C%d" % (last + 1)).WrapText = False


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
        diffs = []
        for n in SNAP:
            old = snap[(n, bd)]
            new = wb.Worksheets(n).Range("A1:AI%d" % len(old)).Value
            d, w = maxdiff(old, new)
            if d > 1e-9:
                diffs.append("%s %.3g@%s" % (n, d, w[:2] if w else w))
        v = t2.Range("B%d:%s%d" % (k["live"], col(5 + k["n"]), k["live"])).Value[0]
        hdr = t2.Range("F%d:%s%d" % (k["live"] - 1, col(5 + k["n"]), k["live"] - 1)).Value[0]
        s = dict(zip(["bundle", "run", "scope", "use"] + [str(h).split("]")[0].strip("[") for h in hdr], v))
        stats[bd] = s
        print("bundle %-6s failing Max: %d | T1/T5 %s | 2030 K %s | changed vs v0.14: %s"
              % (bd, len(bad), ck.Cells(chk, 4).Value, s.get("K"), "; ".join(diffs)))
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
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    wb = None
    try:
        wb = xl.Workbooks.Open(SRC, 0, True)
        xl.Calculation = -4135
        print("snapshot v0.14 ...")
        snap = snapshot(wb)
        old_k = {}
        k = t2_rows(wb)
        t2 = wb.Worksheets("Table2_Industry")
        for r in k["snap"]:
            old_k[t2.Range("B%d" % r).Value] = t2.Range("G%d" % r).Value
        wb.SaveAs(DST)
        write_inputs(wb)
        print("o15/o16 formulas fixed:", fix_npno(wb))
        xl.CalculateFull()
        print("v2 rows refilled:", refill_v2(wb))
        refill_table2(wb, k)
        chk = update_check(wb)
        xl.CalculateFull()
        print("verify ...")
        stats = verify(wb, snap, k, chk)
        new_k = {t2.Range("B%d" % r).Value: t2.Range("G%d" % r).Value for r in k["snap"]}
        txt = ", ".join("%s %.1f (v0.14 %.1f)" % (bd, new_k[bd], old_k[bd]) for bd in BUNDLES[1:])
        print("2030 K:", txt)
        update_notes(wb, chk, txt)
        xl.Calculation = -4105
        xl.CalculateFull()
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
    except Exception:
        for b in list(xl.Workbooks):
            b.Close(False)
        xl.Quit()
        if os.path.exists(DST):
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
