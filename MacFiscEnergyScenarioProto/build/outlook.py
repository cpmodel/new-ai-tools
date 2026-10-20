"""Fiscal outlook and carbon pricing for the Macro-Fiscal-Energy-Scenario Tool.

The Fiscal outlook sheet is the tool's front page: pick a country, set the policy levers (new
taxes, new public investment and other spending, carbon tax and feebate, the fiscal rule), and
read the revenue and expenditure outlook, the balance, debt and growth to 2050. The levers act on
DIGNAD-R's policy scenario (macro_live.py builds the recursion); the baseline has no new measures.

Carbon pricing is PriceFeasibility v0.1 (power sector) in formulas: abatement E and the three
political-feasibility metrics F1 (visible tariff rise), F2 (bottom-40% loss after rebates) and F3
(net payment by coal and oil generators), for any tax and feebate rate. A tax-by-feebate grid on
the Pricing sheet finds the highest-abatement mix that keeps F1-F3 under the limits on the Fiscal
outlook sheet, live, for tax only, feebate only and the hybrid. The chosen mix becomes a revenue
path (tax rate x fossil power CO2 on NMIM's path / NMIM GDP); a feebate raises no revenue. The
revenue is split between rebates to the bottom 40%, green public investment, and the budget.

    python3 outlook.py --check ISO3 ...   print the policy scenario from the Python copy (simulate)
"""
import os, sys
import numpy as np, pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import LineChart, Reference

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
PFD = os.path.join(ROOT, "PriceFeasibility")
MFM = os.path.join(ROOT, "MacroFiscalModels")
for p in (PFD, MFM):
    if p not in sys.path:
        sys.path.insert(0, p)
import price_feasibility as pf                                   # noqa: E402

F = "Arial"; INK = "0E1B2A"; ACC = "C05621"; MUT = "5A6B7B"
HDR = PatternFill("solid", fgColor=INK); YEL = PatternFill("solid", fgColor="FFF2CC")
SEC = PatternFill("solid", fgColor="EAE3F2"); GRN = PatternFill("solid", fgColor="D9EAD3")
GRY = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="C8D0D6"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)


# ---- carbon pricing: PriceFeasibility's country parameters and priors ------------------------
PF_PAR = [("pf_e0", "e0_mt", "Fossil power CO2, latest Ember year (Mt)"),
          ("pf_e_avg", "e_avg", "Average emission factor of generation (t/MWh)"),
          ("pf_e_hi", "e_hi", "Emission factor of coal and oil generation (t/MWh)"),
          ("pf_p_res", "p_res", "Residential tariff (USD/MWh)"),
          ("pf_phi", "phi", "Tariff pass-through"),
          ("pf_s_max", "s_max", "Most substitution abatement by 2030 (share)"),
          ("pf_tau_switch", "tau_switch", "Coal-to-gas switching price (USD/t)"),
          ("pf_b40", "b40", "Electricity budget share, bottom 40%"),
          ("pf_hi_share", "hi_gen_share", "Coal and oil share of generation")]
PF_REF = [("pf_tax_t", "tax_tau_t"), ("pf_tax_E", "tax_E"), ("pf_fee_f", "feebate_tau_f"),
          ("pf_fee_E", "feebate_E"), ("pf_hyb_t", "hybrid_tau_t"), ("pf_hyb_f", "hybrid_tau_f"),
          ("pf_hyb_E", "hybrid_E")]
PR = pf.PRIORS
PF_PRIORS = [("pf_elast", PR["elasticity"], "Price elasticity of electricity demand", "PriceFeasibility PRIORS"),
             ("pf_res_share", PR["res_share"], "Residential share of electricity sales", "PriceFeasibility PRIORS"),
             ("pf_q40", PR["q40"], "Bottom-40% share of residential electricity use", "PriceFeasibility PRIORS"),
             ("pf_wholesale", PR["wholesale_share"], "Wholesale price as a share of the residential tariff",
              "PriceFeasibility PRIORS")]
MAXR = 200                                                      # rate cap (USD/t), as PriceFeasibility
GRID = [0, 1, 2, 3, 4, 5, 7.5, 10, 12.5, 15, 20, 25, 30, 40, 50, 60, 75, 100, 125, 150, 175, 200]
MODES = ["None", "Manual", "Feasible carbon tax", "Feasible feebate", "Feasible hybrid"]

# ---- the levers: (name, default, label, unit/note) -------------------------------------------
# MAIN levers sit on the Dashboard, the rest on Policy settings
MAIN = [
    ("f_tax_step", 0.0, "Revenue: increase each year", "pp of GDP a year"),
    ("f_tax_max", 3.0, "Revenue: total increase, at most", "pp of GDP"),
    ("f_pi", 0.0, "Public investment: increase", "pp of GDP a year"),
    ("f_prog", "On", "Climate investment programme (NMIM budget need)", "On / Off"),
    ("f_cp_mode", "Feasible hybrid", "Carbon pricing, power sector", "see Carbon pricing"),
]
LEVERS = [
    ("New revenue and spending", None, None, None),
    ("f_tax_start", 2026, "New revenue: first year", "year"),
    ("f_tax_cap", 0.5, "Share of new taxes (and carbon tax) that falls on capital", "0 = VAT-like, no effect on investment; 1 = like the base tax"),
    ("f_oth", 0.0, "Other new spending (wages, goods, transfers)", "pp of GDP a year, at full scale"),
    ("f_sp_start", 2026, "New spending: first year", "year"),
    ("f_sp_ramp", 3, "New spending: years to full scale", "years"),
    ("f_rule", 1.0, "Fiscal rule strength", "1 = as calibrated (transfers adjust to debt); 0 = off, debt absorbs everything"),
    ("Carbon pricing (PriceFeasibility)", None, None, None),
    ("f_cp_tax", 0.0, "Manual mode: carbon tax", "USD/tCO2"),
    ("f_cp_fee", 0.0, "Manual mode: feebate rate", "USD/tCO2"),
    ("f_cp_start", 2026, "First year", "year"),
    ("f_cp_phase", 1, "Phase-in", "years in equal steps (also divides F1, the first-year tariff rise)"),
    ("f_F1", pf.DEFAULT_THRESHOLDS["f1"], "Limit F1: first-year rise in the residential tariff", "%"),
    ("f_F2", pf.DEFAULT_THRESHOLDS["f2"], "Limit F2: net loss of the bottom 40%", "% of their consumption"),
    ("f_F3", pf.DEFAULT_THRESHOLDS["f3"], "Limit F3: net payment by coal and oil generators", "% of their revenue"),
    ("f_cp_rebate", 0.0, "Revenue use: rebated to the bottom 40%", "share of carbon tax revenue (lowers F2)"),
    ("f_cp_green", 0.0, "Revenue use: green public investment", "share of carbon tax revenue; the rest goes to the budget"),
]
PS = "'Carbon pricing'"


def pricing_inputs(isos, years):
    """Per country: PriceFeasibility parameters and reference optima, and NMIM's power CO2 (baseline
    and net zero, Mt) and GDP (USD bn) paths. Missing values stay NaN (blank in the workbook)."""
    cal = pd.read_csv(os.path.join(PFD, "output", "calibration_inputs.csv")).set_index("iso3")
    par = pf.country_parameters(cal)
    cm = pd.read_csv(os.path.join(PFD, "output", "country_metrics.csv")).set_index("iso3")
    import nmim_bridge
    m, D = nmim_bridge._nmim()
    yi = [m.YEARS.index(y) for y in years]
    out = pd.DataFrame(index=pd.Index(list(isos), name="iso3"))
    for nm, src, _ in PF_PAR:
        out[nm] = par[src].reindex(out.index)
    for nm, src in PF_REF:
        out[nm] = cm[src].reindex(out.index)
    rows = {}
    for iso in out.index:
        if iso not in D:
            continue
        _, V = m.run_country(D[iso])
        co2 = lambda s: [sum(V[f"mit.con.{f}.pow.{s}"][i] * m.P[f"ef_{f}"] for f in ("coa", "nga", "oil")) for i in yi]
        r = {}
        for y, b, n, g in zip(years, co2(1), co2(2), [V["mit.gdp.lvl"][i] for i in yi]):
            r[f"co2b_{y}"], r[f"co2n_{y}"], r[f"gdpusd_{y}"] = b, n, g
        rows[iso] = r
    return out.join(pd.DataFrame.from_dict(rows, orient="index"))


# ---- formulas ----------------------------------------------------------------------------------
def pf_E(T, Fe):
    """Abatement (share of baseline fossil power CO2) at tax T and feebate Fe (formula strings)."""
    return (f"(1-(1-i_pf_s_max*(1-EXP(-(({T})+({Fe}))/i_pf_tau_switch)))"
            f"*(1-MIN(-pf_elast*q_dp1*({T}),1)))")


def pf_ok(T, Fe):
    return (f"AND(100*q_dp1*({T})/MAX(1,f_cp_phase)<=f_F1+1E-9,100*i_pf_b40*q_dp1*({T})*(1-q_cover)<=f_F2+1E-9,"
            f"q_f3t*({T})+q_f3f*({Fe})<=f_F3+1E-9)")


def recursion_extra(c, pr, on_lever, M, rr):
    """Extra Macro_model columns per scenario (c(var): this row, pr(var): last row). on_lever is
    '0' for the baseline and '1' for the policy scenario."""
    L = on_lever
    return {
        "newtax": f"={L}*${M['dtax']}{rr}*{c('GDP')}",
        "crev": f"={L}*${M['crev']}{rr}*{c('GDP')}",
        "rebate": f"=f_cp_rebate*{c('crev')}",
        "green": f"=f_cp_green*{c('crev')}",
        "newpi": f"={L}*f_pi/100*${M['ramp']}{rr}*{c('GDP')}",
        "newoth": f"={L}*f_oth/100*${M['ramp']}{rr}*{c('GDP')}",
    }


def shared_columns(rows, M, rr, t):
    """Year-level columns: lever paths and the carbon pricing path (policy scenario only). Power CO2
    and GDP come live from NMIM (GI_Model)."""
    import nmim_live
    y = 2024 + t
    gdp = f"=IFERROR({nmim_live.ref(rows, 'mit.gdp.lvl', y)},NA())"
    co2 = lambda sc: f"IFERROR({nmim_live.power_co2(rows, y, sc)},0)"
    if t == 0:
        return {k: "=0" for k in ("dtax", "ramp", "cpT", "cpF", "cpE", "crev")} | {
            "gdpusd": gdp, "co2": f"=IF(p_programme=1,{co2(2)},{co2(1)})"}
    yr = f"$B{rr}"
    ph = f"IF({yr}>=f_cp_start,MIN(1,({yr}-f_cp_start+1)/MAX(1,f_cp_phase)),0)"
    return {
        "dtax": f"=IF({yr}>=f_tax_start,MIN(f_tax_max,f_tax_step*({yr}-f_tax_start+1)),0)/100",
        "ramp": f"=IF({yr}>=f_sp_start,MIN(1,({yr}-f_sp_start+1)/MAX(1,f_sp_ramp)),0)",
        "cpT": f"=cp_T*{ph}", "cpF": f"=cp_F*{ph}",
        "cpE": f"=IFERROR({pf_E(M['cpT'] + str(rr), M['cpF'] + str(rr))},0)",
        "gdpusd": gdp,
        "co2": f"=IF(p_programme=1,{co2(2)},{co2(1)}*(1-{M['cpE']}{rr}))",
        "crev": f'=IFERROR({M["cpT"]}{rr}*{M["co2"]}{rr}/1000/{M["gdpusd"]}{rr},0)',
    }


# ---- sheets ------------------------------------------------------------------------------------
def add_sheets(wb, ctx):
    """Fiscal outlook (front page, country picker, levers, outlook table, charts) and Pricing."""
    M, R0, RL, H0, MC, wm, rng, last_dash = (ctx[k] for k in ("M", "R0", "RL", "H0", "MC", "wm", "rng", "last_dash"))
    at = lambda col, yr: f"Macro_model!${M[col]}${R0 + yr - 2024}"

    # ---- Pricing: scalars, live grid, the three optima -----------------------------------------
    wp = wb.create_sheet("Carbon pricing")
    wp.sheet_view.showGridLines = False
    wp.column_dimensions["A"].width = 2; wp.column_dimensions["B"].width = 44
    wp["B2"] = "Carbon pricing in the power sector: what is feasible, and what it raises"; wp["B2"].font = font(16, True)
    wp["B3"] = ("PriceFeasibility v0.1 in formulas, for the country picked on the Dashboard. A carbon tax raises the "
                "tariff (abatement through substitution and lower demand) and raises revenue; a feebate charges generators "
                "above the average emission rate and pays those below it, so it abates through substitution only, moves "
                "tariffs little and raises nothing. The grid shows abatement (% of baseline fossil power CO2) for each tax "
                "and feebate rate: green where all three limits hold. The mode is set on the Dashboard; limits, phase-in and "
                "rebate share on Policy settings; priors on Macro_assumptions. Illustrative: priors, not "
                "country evidence, set several parameters (see PriceFeasibility/README.md).")
    wp["B3"].font = font(9, italic=True, color=MUT); wp["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    wp.merge_cells("B3:Q3"); wp.row_dimensions[3].height = 66
    wp["B5"] = "Country"; wp["B5"].font = font(10, True)
    wp["C5"] = "=sel_country"; wp["C5"].font = font(12, True, ACC); wp.merge_cells("C5:J5")
    wp["B6"] = ('=IF(ISNA(m_row),"Not covered.",IF(i_pf_e0="","No PriceFeasibility data for this country (no Ember '
                'power data or tariff): carbon pricing is off.",IF(N(i_pf_e0)=0,"No fossil power generation: nothing to price.","")))')
    wp["B6"].font = font(10, True, ACC); wp.merge_cells("B6:Q6")
    r = 8
    wp.cell(r, 2, "Country parameters").font = font(10, True); r += 1
    for nm, _, lab in PF_PAR:
        wp.cell(r, 2, lab).font = font(9)
        x = wp.cell(r, 3, f'=IF(i_{nm}="","",i_{nm})'); x.number_format = "0.000"; x.font = font(9, True)
        wp.merge_cells(f"C{r}:E{r}"); r += 1
    r += 1
    wp.cell(r, 2, "Helpers").font = font(10, True); r += 1
    for nm, lab, f in [("q_dp1", "Tariff rise per USD/t of tax (share)", "=IFERROR(i_pf_phi*i_pf_e_avg/i_pf_p_res,0)"),
                       ("q_need", "Revenue share that fully compensates the bottom 40%", "=IFERROR(i_pf_phi*pf_res_share*pf_q40,0)"),
                       ("q_cover", "Share of the bottom-40% loss the rebate covers", "=IF(q_need>0,MIN(f_cp_rebate/q_need,1),1)"),
                       ("q_f3t", "F3 per USD/t of tax", "=IFERROR(IF(N(i_pf_hi_share)>0,100*i_pf_e_hi*(1-i_pf_phi)/(pf_wholesale*i_pf_p_res),0),0)"),
                       ("q_f3f", "F3 per USD/t of feebate", "=IFERROR(IF(N(i_pf_hi_share)>0,100*MAX(i_pf_e_hi-i_pf_e_avg,0)/(pf_wholesale*i_pf_p_res),0),0)")]:
        wp.cell(r, 2, lab).font = font(9)
        x = wp.cell(r, 3, f); x.number_format = "0.00000"; x.font = font(9)
        wp.merge_cells(f"C{r}:E{r}"); name(wb, nm, f"{PS}!$C${r}"); r += 1

    # grid: tax rates down (rows), feebate rates across (columns); E% shown, score hidden at the right
    r += 1
    G0 = r + 2                                           # first grid row
    n = len(GRID)
    wp.cell(r, 2, "Abatement, % of baseline fossil power CO2 (green: all limits hold)").font = font(10, True)
    wp.cell(r + 1, 2, "Carbon tax (USD/t) down, feebate (USD/t) across").font = font(9, color=MUT)
    C0 = 3                                               # first grid column (C)
    S0 = C0 + n + 2                                      # score grid, hidden
    for j, v in enumerate(GRID):
        for base in (C0, S0):
            x = wp.cell(G0 - 1, base + j, v); x.font = font(8, True, "FFFFFF"); x.fill = HDR
            x.alignment = Alignment(horizontal="center")
    for i, v in enumerate(GRID):
        x = wp.cell(G0 + i, 2, v); x.font = font(8, True); x.alignment = Alignment(horizontal="right")
        wp.cell(G0 + i, S0 - 1, v)
        for j in range(n):
            T = f"$B{G0 + i}"; Fe = f"{get_column_letter(C0 + j)}${G0 - 1}"
            e = wp.cell(G0 + i, C0 + j, f'=IFERROR(100*{pf_E(T, Fe)},"")')
            e.number_format = "0"; e.font = font(8); e.alignment = Alignment(horizontal="center"); e.border = BOX
            wp.cell(G0 + i, S0 + j, f"=IFERROR(IF({pf_ok(T, Fe)},100*{pf_E(T, Fe)}-1E-6*({T}+{Fe}),-1),-1)")
    GL = G0 + n - 1
    cl = lambda c: get_column_letter(c)
    sc = f"${cl(S0)}${G0}:${cl(S0 + n - 1)}${GL}"
    wp.conditional_formatting.add(f"{cl(C0)}{G0}:{cl(C0 + n - 1)}{GL}",
                                  FormulaRule(formula=[f"{cl(S0)}{G0}>=0"], fill=GRN))
    # the optimum, exactly: E rises with both rates; F1 and F2 bound the tax alone, F3 is linear in
    # both. Tax only: the tightest bound. Feebate only: the F3 bound. Hybrid: the best point on the
    # F3 line for taxes from 0 to the tax bound (NS steps), the feebate taking the rest of F3.
    NS = 80
    TB, FB = S0 + n + 1, S0 + n + 2
    wp.cell(G0 - 2, TB, f"=MIN({MAXR},IFERROR(f_F1*MAX(1,f_cp_phase)/(100*q_dp1),{MAXR}),"
                        f"IFERROR(f_F2/(100*i_pf_b40*q_dp1*(1-q_cover)),{MAXR}),IFERROR(f_F3/q_f3t,{MAXR}))")
    name(wb, "q_Tb", f"{PS}!${cl(TB)}${G0 - 2}")
    for k in range(NS + 1):
        rr = G0 + k
        T = f"{cl(TB)}{rr}"; Fe = f"{cl(FB)}{rr}"
        wp.cell(rr, TB, f"=q_Tb*{k}/{NS}")
        wp.cell(rr, FB, f"=MAX(0,MIN({MAXR},IFERROR((f_F3-q_f3t*{T})/q_f3f,{MAXR})))")
        wp.cell(rr, FB + 1, f"=IFERROR(100*{pf_E(T, Fe)}-1E-6*({T}+{Fe}),-1)")
    hs = lambda c: f"${cl(c)}${G0}:${cl(c)}${G0 + NS}"
    for c in list(range(S0 - 1, S0 + n)) + [TB, FB, FB + 1]:
        wp.column_dimensions[cl(c)].hidden = True
    for c in range(C0, C0 + n):
        wp.column_dimensions[cl(c)].width = 5.2

    # results: the three optima, live; the Python reference; what is in use
    r = GL + 2
    wp.cell(r, 2, "Best feasible mix under the limits").font = font(10, True)
    for j, h in enumerate(["Tax USD/t", "Feebate USD/t", "Abatement %", "F1 %", "F2 %", "F3 %", "Python ref.: tax", "feebate", "abatement"]):
        x = wp.cell(r, 3 + j, h); x.font = font(8, True, "FFFFFF"); x.fill = HDR
        x.alignment = Alignment(horizontal="center", wrap_text=True)
    wp.row_dimensions[r].height = 28
    r += 1
    best = lambda c: f"INDEX({hs(c)},MATCH(MAX({hs(FB + 1)}),{hs(FB + 1)},0))"
    opt = [("Carbon tax only", "o_tax", "=q_Tb", "=0", ("pf_tax_t", None, "pf_tax_E")),
           ("Feebate only", "o_fee", "=0", f"=MIN({MAXR},IFERROR(f_F3/q_f3f,{MAXR}))", (None, "pf_fee_f", "pf_fee_E")),
           ("Hybrid (tax + feebate)", "o_hyb", f"={best(TB)}", f"={best(FB)}", ("pf_hyb_t", "pf_hyb_f", "pf_hyb_E"))]
    for lab, nm, ft, ff, ref in opt:
        wp.cell(r, 2, lab).font = font(9, True)
        _row_metrics(wp, r, ft, ff, nm, wb)
        for j, k in enumerate(ref):
            x = wp.cell(r, 9 + j, f'=IF(i_{k}="","",i_{k})' if k else 0)
            x.font = font(9, color=MUT); x.number_format = "0.0"; x.alignment = Alignment(horizontal="center")
        r += 1
    wp.cell(r, 2, "In use (mode on the Dashboard)").font = font(9, True, ACC)
    _row_metrics(wp, r, '=CHOOSE(cp_mode,0,f_cp_tax,o_tax_T,0,o_hyb_T)', '=CHOOSE(cp_mode,0,f_cp_fee,0,o_fee_F,o_hyb_F)', "cp", wb)
    for c in range(3, 9):
        wp.cell(r, c).fill = SEC
    r += 2
    for t_ in ["Best mix: tax only and feebate only are solved exactly (every limit is linear in the rate); the hybrid searches "
               "the F3 limit line in 80 tax steps. Rates are capped at 200 USD/t. Python ref.: PriceFeasibility's own optimiser at "
               "its default limits, no rebate and no phase-in, on its rate grid (1 USD/t steps above 20).",
               "Revenue: tax rate x fossil power CO2 x phase-in, over NMIM GDP. With the climate programme on, CO2 follows "
               "NMIM's net zero path, which already includes the abatement; with it off, NMIM's baseline path less the "
               "abatement above. Abatement is PriceFeasibility's 2030-style response, applied from the year the rate applies."]:
        wp.cell(r, 2, t_).font = font(8, italic=True, color=MUT)
        wp.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top"); wp.merge_cells(f"B{r}:Q{r}")
        wp.row_dimensions[r].height = 30; r += 1
    wp["B5"].font = font(10, True)

    # ---- Policy settings: the secondary levers -----------------------------------------------
    wsx = wb.create_sheet("Policy settings")
    wsx.sheet_view.showGridLines = False
    for col, w in zip("ABCD", [2, 52, 14, 70]):
        wsx.column_dimensions[col].width = w
    wsx["B2"] = "Policy settings: the levers behind the main ones"; wsx["B2"].font = font(16, True)
    wsx["B3"] = ("The main levers (new revenue, public investment, the climate programme, the carbon pricing mode) are on the "
                 "Dashboard. These refine them. Model assumptions: Macro_assumptions (macro model, pricing priors) and "
                 "GI_Inputs (investment needs).")
    wsx["B3"].font = font(9, italic=True, color=MUT); wsx["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    wsx.merge_cells("B3:D3"); wsx.row_dimensions[3].height = 30
    r = 5
    for key, dflt, lab, unit in LEVERS:
        if dflt is None:
            for c in "BCD":
                wsx[f"{c}{r}"].fill = SEC
            wsx[f"B{r}"] = key; wsx[f"B{r}"].font = font(10, True); r += 1
            continue
        wsx[f"B{r}"] = lab; wsx[f"B{r}"].font = font(9); wsx[f"B{r}"].border = BOX
        x = wsx[f"C{r}"]; x.value = dflt; x.fill = YEL; x.font = font(10, True, "0000FF"); x.border = BOX
        x.alignment = Alignment(horizontal="center")
        wsx[f"D{r}"] = unit; wsx[f"D{r}"].font = font(8, italic=True, color=MUT)
        name(wb, key, f"'Policy settings'!$C${r}")
        r += 1
    name(wb, "cp_mode", f"{PS}!$Z$1")
    wp["Z1"] = '=IFERROR(MATCH(f_cp_mode,{"' + '","'.join(MODES) + '"},0),1)'
    wp.column_dimensions["Z"].hidden = True

    # ---- Fiscal outlook: the table and charts -------------------------------------------------
    wo = wb.create_sheet("Fiscal outlook")
    wo.sheet_view.showGridLines = False
    YRS = [2024, 2025, 2026, 2027, 2028, 2029, 2030, 2035, 2040, 2045, 2050]
    YC = [get_column_letter(3 + j) for j in range(len(YRS))]           # C..M
    wo.column_dimensions["A"].width = 2; wo.column_dimensions["B"].width = 50
    for c in YC:
        wo.column_dimensions[c].width = 8
    wo["B2"] = "Fiscal outlook: revenue, spending, balance and debt to 2050"; wo["B2"].font = font(16, True)
    wo["B3"] = ("The Dashboard's scenario item by item, computed live by DIGNAD-R (Macro & debt). The baseline holds the 2024 "
                "position except that fossil rents fall: a reference point, not a forecast; growth is shown against trend.")
    wo["B3"].font = font(9, italic=True, color=MUT); wo["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    wo.merge_cells("B3:M3"); wo.row_dimensions[3].height = 26
    wo["B5"] = "=sel_country"; wo["B5"].font = font(12, True, ACC)
    wo["C5"] = ('="Revenue +"&TEXT(f_tax_step,"0.0")&"pp of GDP a year (to "&TEXT(f_tax_max,"0.0")&"pp); public investment +"'
                '&TEXT(f_pi,"0.0")&"pp; climate programme "&f_prog&"; carbon pricing: "&f_cp_mode&" (tax "&TEXT(cp_T,"0")'
                '&", feebate "&TEXT(cp_F,"0")&" USD/t, "&IF(cp_T+cp_F=0,"none",IF(AND(cp_F1<=f_F1+1E-9,cp_F2<=f_F2+1E-9,'
                'cp_F3<=f_F3+1E-9),"within the limits","OVER a limit"))&")."')
    wo["C5"].font = font(9, True, MUT); wo.merge_cells("C5:M5")
    wo["C5"].alignment = Alignment(wrap_text=True, vertical="top"); wo.row_dimensions[5].height = 26
    r = 7
    # outlook table
    r += 1
    for c in ["B"] + YC:
        wo[f"{c}{r}"].fill = HDR
    wo[f"B{r}"] = "Policy scenario (% of GDP unless noted)"; wo[f"B{r}"].font = font(10, True, "FFFFFF")
    for c, y in zip(YC, YRS):
        x = wo[f"{c}{r}"]; x.value = y; x.font = font(9, True, "FFFFFF"); x.alignment = Alignment(horizontal="center")
    r += 1
    g = lambda X, y: at(f"{X}_GDP", y)
    sh = lambda expr_fn, X="P": (lambda y: f"=({expr_fn(y)})/{g(X, y)}*100")
    rows = [
        ("Growth", None),
        ("GDP growth above trend (pp)", lambda y: f"=({g('P', y)}/{g('P', y - 1)}-1)*100" if y > 2024 else None),
        ("GDP vs the no-new-measures baseline (%)", lambda y: f"=({g('P', y)}/{g('B', y)}-1)*100"),
        ("Revenue", None),
        ("Tax revenue, current policy", sh(lambda y: f"c_tax*{at('P_Y', y)}")),
        ("New taxes", sh(lambda y: at("P_newtax", y))),
        ("Carbon tax revenue (power)", sh(lambda y: at("P_crev", y))),
        ("Resource revenue (government share of fossil rents)", sh(lambda y: f"p_gov_take*{at('rents', y)}")),
        ("Grants", sh(lambda y: at("P_grants", y))),
        ("Total revenue", sh(lambda y: at("P_revenue", y)), "bold"),
        ("Expenditure", None),
        ("Government consumption", sh(lambda y: "c_g")),
        ("Public investment, current policy", sh(lambda y: "c_ig")),
        ("New public investment", sh(lambda y: at("P_newpi", y))),
        ("Climate programme: capital (green public investment)", sh(lambda y: f"p_programme*{at('cap', y)}")),
        ("Climate programme: land use and non-CO2 spending", sh(lambda y: f"p_programme*{at('spend', y)}")),
        ("Carbon revenue: green public investment", sh(lambda y: at("P_green", y))),
        ("Carbon revenue: rebates to the bottom 40%", sh(lambda y: at("P_rebate", y))),
        ("Other new spending", sh(lambda y: at("P_newoth", y))),
        ("Transfers (adjust to debt under the fiscal rule)", sh(lambda y: at("P_tr", y))),
        ("Interest", sh(lambda y: at("P_interest", y))),
        ("Total expenditure", lambda y: f"={at('o_exp_P', y)}", "bold"),
        ("Balance and debt", None),
        ("Primary balance", sh(lambda y: f"-{at('P_prim', y)}")),
        ("Overall balance", lambda y: f"={at('o_bal_P', y)}", "bold"),
        ("Public debt", lambda y: f"={at('P_ratio', y)}*100", "bold"),
        ("Memo: green public investment, all (programme capital + carbon revenue)",
         sh(lambda y: f"p_programme*{at('cap', y)}+{at('P_green', y)}")),
        ("Memo: fossil power CO2 (Mt)", lambda y: f"={at('co2', y)}"),
        ("Memo: carbon tax in force (USD/t)", lambda y: f"={at('cpT', y)}"),
        ("Memo: feebate rate in force (USD/t)", lambda y: f"={at('cpF', y)}"),
        ("Baseline: no new measures, no programme", None),
        ("Total revenue", sh(lambda y: at("B_revenue", y), "B")),
        ("Total expenditure", lambda y: f"={at('o_exp_B', y)}"),
        ("Overall balance", lambda y: f"={at('o_bal_B', y)}"),
        ("Public debt", lambda y: f"={at('B_ratio', y)}*100"),
        ("GDP growth above trend (pp)", lambda y: f"=({g('B', y)}/{g('B', y - 1)}-1)*100" if y > 2024 else None),
        ("Policy less baseline", None),
        ("Public debt (pp of GDP)", lambda y: f"={at('debt_diff_pp', y)}", "bold"),
        ("Overall balance (pp of GDP)", lambda y: f"={at('o_bal_P', y)}-{at('o_bal_B', y)}"),
    ]
    T0 = r
    for row in rows:
        lab, fn = row[0], row[1]
        if fn is None:
            wo[f"B{r}"] = lab; wo[f"B{r}"].font = font(9, True, ACC); r += 1
            continue
        b = len(row) > 2
        wo[f"B{r}"] = lab; wo[f"B{r}"].font = font(9, b); wo[f"B{r}"].border = BOX
        for c, y in zip(YC, YRS):
            f = fn(y)
            x = wo[f"{c}{r}"]
            x.value = f'=IF(ISNA(m_row),"",IFERROR({f[1:]},""))' if f else ""
            x.number_format = "0.0" if "USD/t" not in lab and "Mt" not in lab else "0"
            if "trend" in lab or "Overall" in lab or "Primary" in lab or "pp" in lab or "GDP vs" in lab:
                x.number_format = "0.0;[Red]-0.0"
            x.font = font(9, b); x.border = BOX; x.alignment = Alignment(horizontal="right")
            if b:
                x.fill = GRY
        r += 1
    r += 1
    for t_ in [
        "How the levers work. New taxes rise by the step each year from the first year up to the ceiling; the share on capital "
        "lowers desired private capital as the base tax does (the rest is VAT-like and only cuts household income). New public "
        "investment builds public capital at the country's investment efficiency (PEFA PI-11), which raises output by the return "
        "on public capital; other new spending does not. Spending ramps up over the years to full scale. Everything not paid by "
        "revenue or grants is borrowed at the commercial rate (the climate programme also gets grants and concessional loans, "
        "Macro_assumptions).",
        "The fiscal rule. Transfers fall as debt rises above its 2024 ratio (speed from the integrated run). It offsets part of any "
        "debt-financed package over time: compare the transfers line with the baseline, or set the rule strength to 0.",
        "Carbon pricing. A carbon tax adds revenue and, through the share on capital, a small growth cost; a feebate abates without "
        "revenue or growth cost in this model. Feasible modes pick the highest-abatement rate that keeps F1-F3 under the limits "
        "(Carbon pricing sheet). The integrated run (Macro & debt) is pre-computed with PriceFeasibility's default hybrid and "
        "does not respond to these levers.",
    ]:
        wo[f"B{r}"] = t_; wo[f"B{r}"].font = font(8, italic=True, color=MUT)
        wo[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top"); wo.merge_cells(f"B{r}:M{r}")
        wo.row_dimensions[r].height = 44; r += 1
    wo.freeze_panes = "C7"

    import xlcharts
    cat = (2, R0 + 1, RL)                  # 2025-2050 on a numeric year axis
    S = lambda cols_: [(MC.index(c) + 2, H0, RL, role) for c, role in cols_]
    for k, (title, ser, fmt) in enumerate([
            ("Revenue and expenditure, % of GDP", S([("o_rev_P", "policy"), ("o_exp_P", "run"), ("o_rev_B", "baseline"),
                                                     ("o_exp_B", "run_base")]), "0"),
            ("Public debt, % of GDP", S([("ch_debt_P", "policy"), ("ch_debt_B", "baseline"), ("ch_debt_py", "run")]), "0"),
            ("Overall balance, % of GDP", S([("o_bal_P", "policy"), ("o_bal_B", "baseline")]), "0.0"),
            ("New measures, % of GDP", S([("o_newtax", "policy"), ("o_crev", "gold"), ("o_grn", "green"),
                                          ("o_newsp", "purple")]), "0.0"),
            ("Carbon price in force, USD/t", S([("ch_cpT", "run"), ("ch_cpF", "alt")]), "0")]):
        wo.add_chart(xlcharts.line(wm, cat, ser, title, yfmt=fmt, width=12), f"{'OW'[k % 2]}{2 + 15 * (k // 2)}")
    _dashboard(wb, ctx)
    return wo


def _dashboard(wb, ctx):
    """The summary: country selector, the main levers, headline results and the key charts."""
    import xlcharts
    M, R0, RL, H0, MC, wm, rng, last_dash = (ctx[k] for k in ("M", "R0", "RL", "H0", "MC", "wm", "rng", "last_dash"))
    at = lambda col, yr: f"Macro_model!${M[col]}${R0 + yr - 2024}"
    wd = wb.create_sheet("Dashboard")
    wd.sheet_view.showGridLines = False
    for col, w in zip("ABCDE", [2, 44, 11, 11, 2]):
        wd.column_dimensions[col].width = w
    wd["B2"] = "Macro-Fiscal-Energy-Scenario Tool"; wd["B2"].font = font(18, True)
    wd["B3"] = ("Pick a country and a scenario; every tab follows. Detail tabs after the Detail ->> divider: Fiscal outlook, "
                "Macro & debt, Green investment, Carbon pricing; more levers on Policy settings. Illustrative calibration.")
    wd["B3"].font = font(9, italic=True, color=MUT); wd["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    wd.merge_cells("B3:D3"); wd.row_dimensions[3].height = 40
    wd["B5"] = "Country"; wd["B5"].font = font(11, True)
    c = wd["C5"]; c.value = "Kenya"; c.font = font(12, True, "0000FF"); c.fill = PatternFill("solid", fgColor="FFFF00")
    c.border = BOX; c.alignment = Alignment(horizontal="center"); wd.merge_cells("C5:D5")
    dv = DataValidation(type="list", formula1=f"=Data_wide!$B$2:$B${last_dash}", allow_blank=False)
    wd.add_data_validation(dv); dv.add("C5")
    name(wb, "sel_country", "Dashboard!$C$5")
    wd["B6"] = ('=IF(ISNA(m_row),"Not covered by the models: indicators only.",INDEX(' + rng("income") +
                ',m_row)&"; integrated run: "&INDEX(' + rng("model") + ',m_row))')
    wd["B6"].font = font(9, True, MUT); wd.merge_cells("B6:D6")

    def band(r, text, years=False):
        for col in "BCD":
            wd[f"{col}{r}"].fill = HDR
        wd[f"B{r}"] = text; wd[f"B{r}"].font = font(10, True, "FFFFFF")
        if years:
            for col, y in (("C", 2030), ("D", 2050)):
                x = wd[f"{col}{r}"]; x.value = y; x.font = font(9, True, "FFFFFF"); x.alignment = Alignment(horizontal="center")
    band(8, "Scenario")
    r = 9
    for key, dflt, lab, unit in MAIN:
        wd[f"B{r}"] = f"{lab} ({unit})" if key not in ("f_prog", "f_cp_mode") else lab
        wd[f"B{r}"].font = font(9); wd[f"B{r}"].border = BOX
        x = wd[f"C{r}"]; x.value = dflt; x.fill = YEL; x.font = font(10, True, "0000FF"); x.border = BOX
        x.alignment = Alignment(horizontal="center")
        if key in ("f_prog", "f_cp_mode"):
            opts = ["On", "Off"] if key == "f_prog" else MODES
            d = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=False)
            wd.add_data_validation(d); d.add(f"C{r}")
            if key == "f_cp_mode":
                wd.merge_cells(f"C{r}:D{r}")
        name(wb, key, f"Dashboard!$C${r}")
        r += 1
    wd["H1"] = '=IF(f_prog="Off",0,1)'; name(wb, "p_programme", "Dashboard!$H$1")
    wd.column_dimensions["H"].hidden = True
    wd[f"B{r}"] = ('="Carbon pricing in use: tax "&TEXT(cp_T,"0")&", feebate "&TEXT(cp_F,"0")&" USD/t, abating "&TEXT(cp_E,"0")'
                   '&"% of power CO2; "&IF(cp_T+cp_F=0,"no pricing",IF(AND(cp_F1<=f_F1+1E-9,cp_F2<=f_F2+1E-9,cp_F3<=f_F3+1E-9),'
                   '"within the feasibility limits","OVER a feasibility limit"))&"."')
    wd[f"B{r}"].font = font(8, italic=True, color=ACC); wd.merge_cells(f"B{r}:D{r}"); r += 2

    band(r, "Headline results", years=True); r += 1
    pg = lambda col, X="P": (lambda y: f"{at(col, y)}/{at(X + '_GDP', y)}*100")
    kpis = [
        ("Public debt, % of GDP", lambda y: f"{at('P_ratio', y)}*100", "0.0", True),
        ("   no new measures", lambda y: f"{at('B_ratio', y)}*100", "0.0", False),
        ("   integrated run, climate programme", lambda y: f"{at('ch_debt_py', y)}", "0.0", False),
        ("Overall balance, % of GDP", lambda y: at("o_bal_P", y), "0.0", True),
        ("Revenue, % of GDP", lambda y: at("o_rev_P", y), "0.0", False),
        ("Expenditure, % of GDP", lambda y: at("o_exp_P", y), "0.0", False),
        ("GDP vs no new measures, %", lambda y: at("gdp_dev_prog", y), "0.00", True),
        ("Green investment need, incremental, % of GDP", lambda y: at("gi_dinv", y), "0.00", True),
        ("   of which budget need (the climate programme)", lambda y: at("gi_bud", y), "0.00", False),
        ("Green public investment in the scenario, % of GDP", lambda y: at("o_grn", y), "0.00", False),
        ("Carbon revenue, % of GDP", lambda y: at("o_crev", y), "0.00", False),
    ]
    for lab, fn, fmt, bold in kpis:
        wd[f"B{r}"] = lab; wd[f"B{r}"].font = font(9, bold); wd[f"B{r}"].border = BOX
        for col, y in (("C", 2030), ("D", 2050)):
            x = wd[f"{col}{r}"]; x.value = f'=IF(ISNA(m_row),"",IFERROR({fn(y)},""))'; x.number_format = fmt
            x.font = font(10, bold); x.border = BOX; x.alignment = Alignment(horizontal="center")
        r += 1
    gi = lambda c: f"INDEX({rng(c)},m_row)"
    wd[f"B{r}"] = "Debt risk, LIC-DSF-lite (before / with programme)"; wd[f"B{r}"].font = font(9, True); wd[f"B{r}"].border = BOX
    wd[f"C{r}"] = f'=IFERROR(IF({gi("dsf_rating_base")}="","not assessed",{gi("dsf_rating_base")}&" / "&{gi("dsf_rating_programme")}),"")'
    wd[f"C{r}"].font = font(10, True, ACC); wd[f"C{r}"].alignment = Alignment(horizontal="center"); wd.merge_cells(f"C{r}:D{r}")
    r += 2
    wd[f"B{r}"] = ("Policy = the scenario above; no new measures = today's position held (fossil rents fall). The integrated run "
                   "is pre-computed with the climate programme and default carbon pricing; it does not respond to the levers.")
    wd[f"B{r}"].font = font(8, italic=True, color=MUT); wd[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    wd.merge_cells(f"B{r}:D{r}"); wd.row_dimensions[r].height = 40

    cat = (2, R0 + 1, RL)                  # 2025-2050 on a numeric year axis
    S = lambda cols_: [(MC.index(c) + 2, H0, RL, role) for c, role in cols_]
    wv, c0, h0, dl = ctx["dsa"]
    charts = [
        xlcharts.line(wm, cat, S([("ch_debt_P", "policy"), ("ch_debt_B", "baseline"), ("ch_debt_py", "run")]),
                      "Public debt, % of GDP", width=12),
        xlcharts.line(wm, cat, S([("o_rev_P", "policy"), ("o_exp_P", "run"), ("o_rev_B", "baseline"), ("o_exp_B", "run_base")]),
                      "Revenue and expenditure, % of GDP", width=12),
        xlcharts.line(wm, cat, S([("ch_gdpdev_P", "policy"), ("ch_gdpdev_py", "run")]),
                      "GDP vs no new measures, %", width=12, yfmt="0.0"),
        xlcharts.line(wm, cat, S([("gi_dinv", "green"), ("gi_bud", "green2"), ("o_grn", "policy")]),
                      "Green investment, % of GDP", width=12, yfmt="0.0"),
        xlcharts.line(wm, cat, S([("ch_cpT", "run"), ("ch_cpF", "alt")]), "Carbon price in force, USD/t", width=12),
        xlcharts.line(wv, (c0, h0 + 1, dl), [(c0 + 1, h0, dl, "baseline"), (c0 + 2, h0, dl, "policy"), (c0 + 3, h0, dl, "limit")],
                      "Debt sustainability: PV of public debt, % of GDP", width=12, skip=1),
    ]
    for k, ch in enumerate(charts):
        wd.add_chart(ch, f"{'FN'[k % 2]}{2 + 14 * (k // 2)}")
    return wd


def _row_metrics(ws, r, ft, ff, nm, wb):
    """Tax, feebate, E, F1, F2, F3 for one mix in row r, columns C..H; names <nm>_T, <nm>_F (and
    cp_E, cp_F1.. for the mix in use)."""
    T, Fe = f"$C${r}", f"$D${r}"
    cells = [ft, ff, f'=IFERROR(100*{pf_E(T, Fe)},0)', f"=100*q_dp1*{T}/MAX(1,f_cp_phase)",
             f"=100*i_pf_b40*q_dp1*{T}*(1-q_cover)", f"=q_f3t*{T}+q_f3f*{Fe}"]
    for j, f in enumerate(cells):
        x = ws.cell(r, 3 + j, f"=IFERROR({f[1:]},0)")
        x.number_format = "0.0" if j != 4 else "0.00"; x.font = font(9, True); x.alignment = Alignment(horizontal="center")
        x.border = BOX
    name(wb, f"{nm}_T", f"{PS}!$C${r}"); name(wb, f"{nm}_F", f"{PS}!$D${r}")
    if nm == "cp":
        for j, k in enumerate(["E", "F1", "F2", "F3"]):
            name(wb, f"cp_{k}", f"{PS}!${get_column_letter(5 + j)}${r}")


def name(wb, n, ref):
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


# ---- Python copy of the policy scenario, to check the workbook ----------------------------------
LEVER_DEFAULTS = {k: v for k, v, *_ in MAIN + LEVERS if v is not None}


def best_mix(p, lv):
    """Tax and feebate in use (USD/t), as the Pricing sheet: exact single-instrument bounds, F3-line search."""
    if any(pd.isna(p[k]) for k in ("pf_phi", "pf_e_avg", "pf_p_res")):
        return 0.0, 0.0
    PV = dict((k, v) for k, v, *_ in PF_PRIORS)
    dp1 = p.pf_phi * p.pf_e_avg / p.pf_p_res
    need = p.pf_phi * PV["pf_res_share"] * PV["pf_q40"]
    cover = min(lv["f_cp_rebate"] / need, 1) if need > 0 else 1
    hi = p.pf_hi_share > 0
    f3t = 100 * p.pf_e_hi * (1 - p.pf_phi) / (PV["pf_wholesale"] * p.pf_p_res) if hi else 0
    f3f = 100 * max(p.pf_e_hi - p.pf_e_avg, 0) / (PV["pf_wholesale"] * p.pf_p_res) if hi else 0
    div = lambda a, b: a / b if b else MAXR
    Tb = min(MAXR, div(lv["f_F1"] * max(1, lv["f_cp_phase"]), 100 * dp1),
             div(lv["f_F2"], 100 * p.pf_b40 * dp1 * (1 - cover)), div(lv["f_F3"], f3t))
    Fb = min(MAXR, div(lv["f_F3"], f3f))
    E = lambda T, Fe: 1 - (1 - p.pf_s_max * (1 - np.exp(-(T + Fe) / p.pf_tau_switch))) * (1 - min(-PV["pf_elast"] * dp1 * T, 1))
    mode = lv["f_cp_mode"]
    if mode == "None":
        return 0.0, 0.0
    if mode == "Manual":
        return lv["f_cp_tax"], lv["f_cp_fee"]
    if mode == "Feasible carbon tax":
        return Tb, 0.0
    if mode == "Feasible feebate":
        return 0.0, Fb
    Ts = [Tb * k / 80 for k in range(81)]
    Fs = [max(0, min(MAXR, div(lv["f_F3"] - f3t * T, f3f))) for T in Ts]
    sc = [E(T, Fe) * 100 - 1e-6 * (T + Fe) for T, Fe in zip(Ts, Fs)]
    i = int(np.argmax(sc))
    return Ts[i], Fs[i]


def simulate(c, pfr, long, lv=None, prm=None):
    """DIGNAD-R baseline (B) and policy (P) paths for one country, with the levers lv (dict)."""
    import build_dignad_excel as dr
    lv = {**LEVER_DEFAULTS, **(lv or {})}
    PV = {**dr.PV, **{k: v for k, v, *_ in PF_PRIORS}, **(prm or {})}
    k = dr.INCOME.index(c.income)
    pefa = c.pefa_pi11_pim_score
    s = 1 - (0.40 - (pefa - 1) / 3 * 0.27 if pd.notna(pefa) else (0.40, 0.27, 0.27, 0.13)[k])
    Rg = c.return_public_capital_pct / 100; ig = c.public_investment_pct_gdp / 100
    rr = c.fossil_rents_total_pct_gdp / 100 if pd.notna(c.fossil_rents_total_pct_gdp) else 0.0
    res = rr >= PV["p_res_threshold"]; rents0 = rr / (1 - rr) if res else 0.0
    conc = c.ids_multilateral_share_pct / 100 if pd.notna(c.ids_multilateral_share_pct) else (0.70, 0.45, 0.15, 0.0)[k]
    prc = c.ids_private_creditor_rate_pct
    rstar = min(max(prc / 100 - 0.03, 0.03), 0.08) if pd.notna(prc) else (0.04 if k < 2 else 0.03)
    tax = c.tax_revenue_pct_gdp / 100; grants0 = c.grants_pct_gdp / 100
    a, d, dg = PV["p_alpha"], PV["p_delta"], PV["p_delta_g"]
    K0 = a * (1 - tax) / (rstar + d); KG0 = s * ig / dg; psi = Rg * KG0; A = 1 / (KG0 ** psi * K0 ** a)
    GDP0 = 1 + rents0; D0 = c.gross_debt_pct_gdp / 100 * GDP0; Dk0 = conc * D0
    int0 = PV["p_rk"] * Dk0 + rstar * (D0 - Dk0)
    rev0 = tax + PV["p_gov_take"] * rents0 + grants0
    g = max(0.05, rev0 - ig - int0 - PV["p_tr_target"]); trbar = rev0 - g - ig - int0
    lam = c.fiscal_rule_speed if pd.notna(c.fiscal_rule_speed) else (PV["p_lambda_f_res"] if res else PV["p_lambda_f"])
    lam *= lv["f_rule"]
    import nmim_bridge                                   # the climate programme, live from NMIM as in GI_Model
    pg = nmim_bridge.programme(c.iso3)
    cap = dict(zip(pg["years"].tolist(), pg["pub_cap"])); spd = dict(zip(pg["years"].tolist(), pg["pub_spend"]))
    mix = best_mix(pfr, lv); cT, cF = mix[0], mix[1]
    PVp = PV
    def cpE(T, Fe):
        if any(pd.isna(pfr[k_]) for k_ in ("pf_phi", "pf_e_avg", "pf_p_res", "pf_s_max", "pf_tau_switch")):
            return 0.0
        dp1 = pfr.pf_phi * pfr.pf_e_avg / pfr.pf_p_res
        return 1 - (1 - pfr.pf_s_max * (1 - np.exp(-(T + Fe) / pfr.pf_tau_switch))) * (1 - min(-PVp["pf_elast"] * dp1 * T, 1))
    out = {}
    prog = 1 if lv["f_prog"] == "On" else 0
    for X, on, L in (("B", 0, 0), ("P", prog, 1)):
        KG, K, D, Dk, rc, GDPp = KG0, K0, D0, Dk0, rstar, GDP0
        path = [dict(GDP=GDP0, ratio=D0 / GDP0, revenue=rev0, crev=0.0, tr=trbar)]
        for t in range(1, dr.T):
            yr = dr.Y0 + t
            dy = PV["p_dis_year"]; dis = 1 if dy and yr == dy else 0
            tfp = 1 - PV["p_eps_a"] * PV["p_rho_a"] ** (yr - dy) if dy and yr >= dy else 1
            rents = rents0 * (1 - PV["p_rent_decline"] * min(1, t / (dr.Y1 - dr.Y0)))
            dtax = (min(lv["f_tax_max"], lv["f_tax_step"] * (yr - lv["f_tax_start"] + 1)) if yr >= lv["f_tax_start"] else 0) / 100
            ramp = min(1, (yr - lv["f_sp_start"] + 1) / max(1, lv["f_sp_ramp"])) if yr >= lv["f_sp_start"] else 0
            ph = min(1, (yr - lv["f_cp_start"] + 1) / max(1, lv["f_cp_phase"])) if yr >= lv["f_cp_start"] else 0
            T, Fe = cT * ph, cF * ph
            E = cpE(T, Fe)
            co2 = pfr[f"co2n_{yr}"] if prog == 1 else pfr[f"co2b_{yr}"] * (1 - E)
            crv = T * co2 / 1000 / pfr[f"gdpusd_{yr}"] if pd.notna(co2) and pd.notna(pfr[f"gdpusd_{yr}"]) else 0.0
            Pc, Ps = on * cap.get(yr, 0), on * spd.get(yr, 0)
            KGd, Kd = KG * (1 - PV["p_eps_g"] * dis), K * (1 - PV["p_eps_k"] * dis)
            Y = A * tfp * KGd ** psi * Kd ** a
            GDP = Y + rents
            newtax, crev = L * dtax * GDP, L * crv * GDP
            rebate, green = lv["f_cp_rebate"] * crev, lv["f_cp_green"] * crev
            newpi, newoth = L * lv["f_pi"] / 100 * ramp * GDP, L * lv["f_oth"] / 100 * ramp * GDP
            Kstar = a * ((1 - tax) * Y - lv["f_tax_cap"] * (newtax + crev)) / (rstar + d)
            Kn = Kd + PV["p_lambda_k"] * (Kstar - Kd)
            KGn = (1 - dg) * KGd + s * (ig + Pc + newpi + green)
            grants = grants0 + PV["p_sh_grant"] * (Pc + Ps)
            revenue = tax * Y + newtax + crev + PV["p_gov_take"] * rents + grants
            tr = trbar - lam * (D / GDPp - D0 / GDP0)
            prim = g + ig + Pc + Ps + newpi + newoth + green + rebate + tr - revenue
            interest = PV["p_rk"] * Dk + rc * (D - Dk)
            Dn = D + interest + prim
            Dkn = Dk0 + (Dk - Dk0) * (1 - PV["p_amort"]) + PV["p_sh_conc"] * (Pc + Ps)
            ratio = Dn / GDP
            rc = rstar + PV["p_eta"] * max(0, ratio - D0 / GDP0)
            KG, K, D, Dk, GDPp = KGn, Kn, Dn, Dkn, GDP
            path.append(dict(GDP=GDP, ratio=ratio, revenue=revenue / GDP, crev=crev / GDP, tr=tr / GDP))
        out[X] = path
    return out, (cT, cF)


if __name__ == "__main__" and "--check" in sys.argv:
    import json
    import macro_live
    isos = [a for a in sys.argv[sys.argv.index("--check") + 1:] if not a.startswith("{")]
    lv = json.loads(next((a for a in sys.argv if a.startswith("{")), "{}"))
    wide, long, _ = macro_live.load(isos)
    pfi = pricing_inputs(wide.iso3, macro_live.YEARS)
    for _, c in wide.iterrows():
        res, mix = simulate(c, pfi.loc[c.iso3], long, lv)
        print(c.iso3, "mix", [round(x, 3) for x in mix])
        for t in (6, 26):
            P, B = res["P"][t], res["B"][t]
            print(c.iso3, 2024 + t, f"debt_P={100 * P['ratio']:.6f}", f"debt_B={100 * B['ratio']:.6f}",
                  f"gdp_P_vs_B={100 * (P['GDP'] / B['GDP'] - 1):.6f}", f"rev_P={100 * P['revenue']:.6f}",
                  f"crev_P={100 * P['crev']:.6f}")
