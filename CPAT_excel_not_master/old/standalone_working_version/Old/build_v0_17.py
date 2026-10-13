"""Build CPAT_Industry_Kernel_Egypt_v0.17.xlsx from v0.16: adds sheet CarveOut_Table2 (final Table 2, CBAM carve-out v0.3).

The sheet reproduces egypt/supporting/AdHocRebuild/make_carveout_v0_3.py line by line with live formulas for the
active bundle (Settings!B10), 2030 (column T): original CPAT run results everywhere, CPAT's implied CBAM-block change
removed, block rebuilt from CPAT's fuel-intensity response + kernel output response + kernel process response (charged
bundles only). O on FULL; NOPHASE memo. A stored 6-bundle snapshot (from make_carveout_v0_3.carve) and a Check row
compare live vs stored. No other sheet's formulas change.

Run with Excel installed (from this folder):  python build_v0_17.py
"""
import datetime
import os
import sys

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "Old"))
ADHOC = os.path.join(HERE, "..", "..", "..", "egypt", "supporting", "AdHocRebuild")
sys.path.insert(0, ADHOC)

from build_v0_4 import GREEN, REVIEW, copy_formats  # noqa: E402
import make_carveout_v0_3 as co  # noqa: E402

SRC = os.path.join(HERE, "Old", "CPAT_Industry_Kernel_Egypt_v0.16.xlsx")
DST = os.path.join(HERE, "CPAT_Industry_Kernel_Egypt_v0.17.xlsx")
VER = "v0.17"
SH = "CarveOut_Table2"
MI, T2, MIT = "'Manual inputs'!", "Table2_Industry!", "Mitigation_Industry!"
BUNDLES = co.BUNDLES
# sector: (label, fuel row, process row, production rows, CPAT growth cell)
SECT = (("mch", "Mining & chemicals: ammonia, urea, ammonium nitrate", 374, 385, (280, 282), "D26"),
        ("irn", "Iron & steel", 375, 386, (276, 278), "D24"),
        ("nfm", "Non-ferrous metals: primary aluminium", 376, 387, (283, 283), "D26"),
        ("cem", "Cement (non-metallic minerals)", 378, 389, (279, 279), "D25"))
POL = 399  # scenario 2 (policy) rows = scenario 1 row + 399


def put(ws, r, label, f, unit="", note="", fmt="0.000", inp=False):
    ws.Range("B%d" % r).Value = label
    ws.Range("C%d" % r).Value = unit
    ws.Range("D%d" % r).Formula = f
    ws.Range("D%d" % r).NumberFormat = fmt
    ws.Range("E%d" % r).Value = note
    if inp:
        ws.Range("D%d" % r).Interior.Color = GREEN


def head(ws, r, text):
    ws.Range("B%d" % r).Value = text
    ws.Range("B%d" % r).Font.Bold = True


def write_sheet(wb):
    ws = wb.Worksheets.Add(After=wb.Worksheets("Table2_Industry"))
    ws.Name = SH
    ws.Range("B1").Value = "Table 2 (2030), final: CPAT results with a CBAM carve-out (v0.3, O on FULL). Live for the active bundle"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13
    ws.Range("B2").Value = ("Original CPAT run everywhere; CPAT's implied CBAM-block change (r x block) removed; block rebuilt "
                            "from CPAT's fuel-intensity response (fuel per t), the kernel output response and, where process "
                            "emissions are charged, the kernel process response. Each effect counted once.")
    ws.Range("B3").Value = ("Set the bundle in Settings!B10 (1A, 2A, 2B, 3A, 3B, 3C). Year 2030 = column T of the source "
                            "sheets. Green = hard input. Reference implementation: egypt/supporting/AdHocRebuild/"
                            "make_carveout_v0_3.py; method: EGYPT_Methodology_v1.3 section 5.0.")

    head(ws, 5, "A. Inputs (active bundle, 2030)")
    put(ws, 6, "Active bundle", "=%s$D$14" % T2, fmt="@")
    put(ws, 7, "CPAT run", "=%s$D$15" % T2, fmt="@", note="1A/2A = EG1, 2B = EG2, 3A-3C = EG3")
    put(ws, 8, "Energy scope", "=%s$D$16" % T2, fmt="@")
    put(ws, 9, "Process emissions charged (0/1)", "=%s$D$17" % T2, fmt="0")
    put(ws, 10, "Revenue use", "=%s$D$18" % T2, fmt="@")
    put(ws, 11, "Output-based free allocation (0/1)", '=IF(ISNUMBER(SEARCH("output-based",D10)),1,0)', fmt="0",
        note="3B: free allocation leaves no domestic-price credit against CBAM")
    put(ws, 12, "GHG0: CPAT total GHG incl. LULUCF, baseline", "=%s$T$25" % T2, "MtCO2e")
    put(ws, 13, "ENR0: CPAT energy CO2, baseline", "=%s$T$29" % T2, "MtCO2")
    put(ws, 14, "IND0: CPAT industry energy CO2, baseline", "=%s$T$31" % T2, "MtCO2")
    put(ws, 15, "IND1: CPAT industry energy CO2, policy (run)", "=%s$T$32" % T2, "MtCO2")
    put(ws, 16, "kappa: share of industry energy CO2 priced by the run", "=%s$T$54" % T2, "share")
    put(ws, 17, "tau: carbon price, active bundle", "=%s$T$71" % T2, "$/tCO2")
    put(ws, 18, "EU ETS / CBAM certificate price", "=%s$T$79" % MI, "$/tCO2")
    put(ws, 19, "CBAM factor used (phase-in)", "=%s$T$81" % MI, "share")
    put(ws, 20, "s_int: efficiency share of CPAT's industry fuel response", "=1/3", "share",
        "Methodology section 3", inp=True)
    put(ws, 21, "Run for the fuel-intensity response", '=IF(D7="EG3","EG1",D7)', fmt="@",
        note="EG3 prices only ~54% of industry; the block is fully priced at the same $20/t, so 3A-3C use EG1")
    put(ws, 22, "IND0 of that run", "=INDEX(CPAT_National!$T$6:$T$133,MATCH(%s$G$31&\"|\"&D21,CPAT_National!$B$6:$B$133,0))" % T2, "MtCO2")
    put(ws, 23, "IND1 of that run", "=INDEX(CPAT_National!$T$6:$T$133,MATCH(%s$G$32&\"|\"&D21,CPAT_National!$B$6:$B$133,0))" % T2, "MtCO2")
    put(ws, 24, "CPAT growth 2024-2030, iron & steel energy CO2", "=Emissions_Industry!$T$15/Emissions_Industry!$N$15", "ratio")
    put(ws, 25, "CPAT growth 2024-2030, cement energy CO2", "=Emissions_Industry!$T$27/Emissions_Industry!$N$27", "ratio")
    put(ws, 26, "CPAT growth 2024-2030, total industry energy CO2", "=IPPU_Industry!$T$11/IPPU_Industry!$N$11", "ratio",
        "CPAT has no mch / nfm energy for Egypt, so they follow total industry")
    put(ws, 27, "Kernel v0.16 O, NOPHASE convention (stored input, table G6:H12)", "=INDEX($H$7:$H$12,MATCH(D6,$G$7:$G$12,0))", "%")
    ws.Range("G6").Value = "Bundle"
    ws.Range("H6").Value = "v0.16 O NOPHASE, %"
    ws.Range("G6:H6").Font.Bold = True
    for i, bd in enumerate(BUNDLES):
        ws.Range("G%d" % (7 + i)).Value = bd
        ws.Range("H%d" % (7 + i)).Value = co.O[bd]
        ws.Range("H%d" % (7 + i)).Interior.Color = GREEN

    head(ws, 29, "B. Rates")
    put(ws, 30, "r: CPAT industry energy CO2 change of the run", "=D15/D14-1", "share",
        "CPAT scales all industry, incl. IPPU, by r: r x block = CPAT's implied block change", fmt="0.0000")
    put(ws, 31, "i_f: CPAT fuel-intensity response = (IND1/IND0)^s_int - 1", "=(D23/D22)^D20-1", "share", fmt="0.0000")
    put(ws, 32, "O credit factor, FULL = (CBF x P_EU - tau)/(CBF x P_EU); 3B = 1", "=IF(D11=1,1,(D19*D18-D17)/(D19*D18))", "share", fmt="0.0000")
    put(ws, 33, "O credit factor, NOPHASE memo = (P_EU - tau)/P_EU; 3B = 1", "=IF(D11=1,1,(D18-D17)/D18)", "share", fmt="0.0000")

    head(ws, 35, "C. CBAM block by sector, 2030 (MtCO2; output in kt; Mitigation_Industry scenario 1 = baseline, +399 = policy)")
    hdr = ("Sector", "Description", "Fuel row", "Process row", "Base 2024 fuel+proc", "Base fuel", "Base proc",
           "Policy fuel", "Policy proc", "Output base", "Output policy", "Output change u", "CPAT growth",
           "Kernel growth", "Growth factor f", "f x base fuel (Bf)", "f x base proc (Bp)",
           "f x (pol fuel x (1+i_f) - base fuel) (dF)", "f x (pol proc - base proc) (dP)",
           "e0 = base fuel+proc", "e1(i_f) = (pol fuel(1+i_f)+pol proc)/(1+u)", "e1(0)")
    for j, h in enumerate(hdr):
        c = ws.Cells(36, 2 + j)
        c.Value = h
        c.Font.Bold = True
        c.WrapText = True
    r0 = 37
    for i, (s, desc, rf, rp, (pa, pb), g) in enumerate(SECT):
        r = r0 + i
        m = MIT
        vals = [s, desc, rf, rp,
                "=%sN%d+%sN%d" % (m, rf, m, rp), "=%sT%d" % (m, rf), "=%sT%d" % (m, rp),
                "=%sT%d" % (m, rf + POL), "=%sT%d" % (m, rp + POL),
                "=SUM(%sT%d:T%d)" % (m, pa, pb), "=SUM(%sT%d:T%d)" % (m, pa + POL, pb + POL),
                "=L{r}/K{r}-1", "=$" + g[0] + "$" + g[1:], "=(G{r}+H{r})/F{r}", "=N{r}/O{r}",
                "=P{r}*G{r}", "=P{r}*H{r}", "=P{r}*(I{r}*(1+$D$31)-G{r})", "=P{r}*(J{r}-H{r})",
                "=G{r}+H{r}", "=(I{r}*(1+$D$31)+J{r})/(1+M{r})", "=(I{r}+J{r})/(1+M{r})"]
        for j, v in enumerate(vals):
            c = ws.Cells(r, 2 + j)
            if isinstance(v, str) and v.startswith("="):
                c.Formula = v.format(r=r)
                c.NumberFormat = "0.0000"
            else:
                c.Value = v
    rt = r0 + len(SECT)
    ws.Range("B%d" % rt).Value = "Total"
    ws.Range("B%d" % rt).Font.Bold = True
    for cl in "FGHIJKLQRSTUVW":
        ws.Range("%s%d" % (cl, rt)).Formula = "=SUM(%s%d:%s%d)" % (cl, r0, cl, rt - 1)
        ws.Range("%s%d" % (cl, rt)).NumberFormat = "0.0000"

    head(ws, 43, "D. Carve-out")
    put(ws, 44, "Bf: block fuel CO2, baseline (growth-adjusted)", "=Q%d" % rt, "MtCO2")
    put(ws, 45, "Bp: block process CO2, baseline (growth-adjusted)", "=R%d" % rt, "MtCO2")
    put(ws, 46, "dF: new block fuel change (output x CPAT fuel intensity)", "=S%d" % rt, "MtCO2")
    put(ws, 47, "dP: new block process change (kernel)", "=T%d" % rt, "MtCO2")
    put(ws, 48, "CPAT's implied block change = r x (Bf + Bp)", "=D30*(D44+D45)", "MtCO2")
    put(ws, 49, "adj_f = dF - r x Bf", "=D46-D30*D44", "MtCO2")
    put(ws, 50, "adj_p = dP - r x Bp", "=D47-D30*D45", "MtCO2")
    put(ws, 51, "dN: fuel-intensity term in CBAM intensity", "=100*(V%d-W%d)/U%d" % (rt, rt, rt), "pts")
    put(ws, 52, "Check: kernel N rebuilt from this table - Table2 N", "=100*(W%d/U%d-1)-%s$T$87" % (rt, rt, T2), "pts",
        "small (<0.2 pt): unscaled block, output rows", fmt="0.000")
    put(ws, 54, "CPAT dGHG of the run", "=%s$T$49" % T2, "MtCO2e")
    put(ws, 55, "CPAT carbon-tax receipts of the run", "=%s$T$47" % T2, "$bn")
    put(ws, 56, "CPAT deaths avoided of the run", "=%s$T$48" % T2, "deaths", fmt="0")
    put(ws, 57, "CPAT dEnergy CO2 of the run", "=%s$T$51" % T2, "MtCO2")
    put(ws, 58, "Kernel block process revenue", "=%s$T$73" % T2, "$bn")
    put(ws, 59, "Kernel output-based rebate (3B)", "=%s$T$74" % T2, "$bn")
    put(ws, 60, "Kernel abatement fund (3C)", "=%s$T$75" % T2, "$bn")
    put(ws, 61, "Kernel N (CBAM intensity)", "=%s$T$87" % T2, "%")
    put(ws, 62, "Kernel O, FULL", "=%s$T$88" % T2, "%")
    put(ws, 63, "Kernel M (CBAM coverage)", "=%s$T$86" % T2, "%")

    head(ws, 65, "E. Table 2 columns, active bundle")
    rows = [
        ("J", "[J] Coverage, % of GHG", '=100*(IF(D8="All sectors",D13,D16*D14)+IF(D9=1,D45,0))/D12', "%", "0",
         "All sectors: energy CO2; Industry only: kappa x IND0; plus block process where charged"),
        ("K", "[K] Emission cut, MtCO2e", "=D54+D49+D50", "MtCO2e", "0.0", "CPAT dGHG + adj_f + adj_p"),
        ("L", "[L] Cut, % of GHG", "=100*D67/D12", "%", "0.0", ""),
        ("M", "[M] CBAM coverage, %", "=D63", "%", "0", "kernel"),
        ("N", "[N] CBAM intensity, %", "=D61+D51", "%", "0.0", "kernel N + dN"),
        ("O", "[O] CBAM obligations per t, % (FULL)", "=D62+D32*D51", "%", "0.0", "kernel O FULL + credit x dN"),
        ("Onp", "memo: [O] NOPHASE", "=D27+D33*D51", "%", "0.0", "v0.16 NOPHASE O + credit x dN"),
        ("P", "[P] Net revenue, $bn", "=D55+D17*D49/1000+D58-D59-D60", "$bn", "0.0",
         "CPAT receipts + tau x adj_f + process fees - rebate - fund"),
        ("Q", "[Q] Deaths avoided", "=D56*(D57+D49)/D57", "deaths", "0", "CPAT deaths x (dEnergy + adj_f)/dEnergy"),
        ("T", "CBAM block emissions change, %", "=100*(D46+D47)/(D44+D45)", "%", "0.0", ""),
    ]
    live = {}
    for i, (k, lab, f, u, fm, note) in enumerate(rows):
        r = 66 + i
        put(ws, r, lab, f, u, note, fmt=fm)
        ws.Range("D%d" % r).Interior.Color = REVIEW
        live[k] = r

    head(ws, 78, "F. Stored 2030 snapshot, all bundles (make_carveout_v0_3.py; = EGYPT_CarveOut_Table2_v0.3)")
    keys = [k for k, *_ in rows]
    ws.Range("B79").Value = "Bundle"
    for j, k in enumerate(keys):
        ws.Cells(79, 3 + j).Value = k
    ws.Range("B79:M79").Font.Bold = True
    _, R = co.carve()
    t2 = wb.Worksheets("Table2_Industry")
    m_snap = {t2.Range("B%d" % r).Value: t2.Range("I%d" % r).Value for r in range(103, 109)}  # v0.16 stored M
    for i, bd in enumerate(BUNDLES):
        r = 80 + i
        ws.Range("B%d" % r).Value = bd
        for j, k in enumerate(keys):
            v = m_snap[bd] if k == "M" else R[bd][k]
            ws.Cells(r, 3 + j).Value = v
            ws.Cells(r, 3 + j).NumberFormat = "0.0"
    ws.Range("B86").Value = "live"
    ws.Range("B87").Value = "live - stored (active bundle)"
    for j, k in enumerate(keys):
        cl = ws.Cells(86, 3 + j).Address.split("$")[1]
        ws.Cells(86, 3 + j).Formula = "=D%d" % live[k]
        ws.Cells(87, 3 + j).Formula = "=IFERROR(%s86-INDEX(%s$80:%s$85,MATCH($D$6,$B$80:$B$85,0)),\"n/a\")" % (cl, cl, cl)
        ws.Cells(86, 3 + j).NumberFormat = "0.0"
        ws.Cells(87, 3 + j).NumberFormat = "0.000"
    ws.Range("B88").Value = "Max |live - stored| (< 0.01 expected: kappa 0.537 rounded in the script)"
    ws.Range("D88").Formula = '=IFERROR(MAX(INDEX(ABS(C87:L87),0)),"n/a")'
    ws.Range("D88").Interior.Color = REVIEW
    ws.Range("B90").Value = ("Not re-solved: 3B rebate and 3C fund (kernel v0.16 values; block fuel-payment change < $0.05bn). "
                             "3B published deaths (345) not traceable. EG3 prices ~54% of industry as CPAT ran it.")
    ws.Columns("B").ColumnWidth = 58
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 14
    ws.Range("E6:E90").WrapText = False
    ws.Range("F36:W36").ColumnWidth = 12
    ws.Rows(36).RowHeight = 60
    return live


def add_check(wb):
    ws = wb.Worksheets("Check")
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    r0 = last + 3
    ws.Range("A%d" % r0).Value = "v0.17: CBAM carve-out (sheet %s)" % SH
    ws.Range("A%d" % r0).Font.Bold = True
    ws.Range("A%d" % (r0 + 1)).Value = "Max |live - stored carve-out Table 2| (active bundle; n/a for LEGACY)"
    ws.Range("D%d" % (r0 + 1)).Formula = "=%s!D88" % SH
    ws.Range("D%d" % (r0 + 1)).Interior.Color = REVIEW
    ws.Range("A%d" % (r0 + 2)).Value = "Kernel N rebuilt from the carve-out block table - Table2 N (pts; < 0.2 expected)"
    ws.Range("D%d" % (r0 + 2)).Formula = "=IFERROR(%s!D52,\"n/a\")" % SH
    return r0 + 1


def update_notes(wb, chk, stats):
    t2 = wb.Worksheets("Table2_Industry")
    t2.Range("B112").Value = ("FINAL Table 2 = CBAM carve-out v0.3 (sheet %s): CPAT runs + block rebuilt; O on FULL. "
                              "Section E above is the prototype (kernel-composition) Table 2, kept for reference." % SH)
    t2.Range("B112").Font.Bold = True
    ws = wb.Worksheets("Settings")
    ws.Range("A1").Value = "CPAT industry kernel - Egypt prototype " + VER
    ws.Range("C25").Value = str(ws.Range("C25").Value) + " %s: carve-out section (row %d)." % (VER, chk)
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    copy_formats(ws.Rows(last), ws.Rows(last + 1))
    wb.Application.CutCopyMode = False
    ws.Range("A%d" % (last + 1)).Value = VER
    ws.Range("B%d" % (last + 1)).Value = "'" + datetime.date.today().isoformat()
    ws.Range("C%d" % (last + 1)).Value = (
        "Final Table 2: sheet %s = CBAM carve-out v0.3 in live formulas (active bundle) + stored 6-bundle snapshot; "
        "O on FULL, NOPHASE memo. No other formulas changed. Check row %d. 2030 K: %s." % (SH, chk, stats))
    ws.Range("C%d" % (last + 1)).WrapText = False


def main():
    if os.path.exists(DST):
        raise SystemExit("%s exists; never overwrite a saved version" % DST)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3
    try:
        wb = xl.Workbooks.Open(SRC, 0, True)
        wb.SaveAs(DST)
        live = write_sheet(wb)
        chk = add_check(wb)
        st = wb.Worksheets("Settings")
        b0 = st.Range("B10").Value
        co_ws = wb.Worksheets(SH)
        worst, ks = 0.0, []
        for bd in BUNDLES:
            st.Range("B10").Value = bd
            xl.CalculateFull()
            d = co_ws.Range("D88").Value
            nchk = co_ws.Range("D52").Value
            print("%s live K %.2f P %.2f Q %.0f J %.1f N %.2f O %.2f | max|diff| %s | Ncheck %.3f" % (
                bd, *(co_ws.Range("D%d" % live[k]).Value for k in "KPQJNO"), d, nchk))
            worst = max(worst, d if isinstance(d, float) else 99)
            ks.append("%s %.1f" % (bd, co_ws.Range("D%d" % live["K"]).Value))
        st.Range("B10").Value = b0
        update_notes(wb, chk, ", ".join(ks))
        xl.CalculateFull()
        if worst > 0.05:
            raise SystemExit("carve-out live vs stored differs by %.3f" % worst)
        wb.Worksheets("Settings").Activate()
        wb.Save()
        wb.Close(False)
        print("worst |live - stored| %.4f" % worst)
    except BaseException:
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
