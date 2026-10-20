"""Live simple macro model for the Macro-Fiscal-Energy-Scenario Tool: DIGNAD-R for the country picked on the
Dashboard, in formulas.

DIGNAD-R is the recursive, formula-only DIGNAD-lite / DIGNAR-lite of
MacroFiscalModels/ExcelDIGNAD/build_dignad_excel.py (equations in ExcelDIGNAD/README.md). Its
assumptions, augmented-input defaults, input loader and Python copy of the equations are imported
from there, so the two stay one model; only the layout differs: DIGNAD-R Excel computes 12
countries at once, this sheet computes the selected one, for any of the integrated run's countries.

Sheets added to the built workbook:
  Macro              view: the integrated run (pre-computed, Python) beside the live simple model
  Macro_assumptions  named yellow cells p_... (change them; everything recalculates)
  Macro_model        calibration and the 2024-2050 recursion, baseline (B) and programme (P)
  Macro_inputs       one row per country: dashboard inputs, augmented defaults, NMIM programme and
                     pricing paths, and the integrated run's GDP and debt paths

    python3 macro_live.py --check ISO3 ...   recompute the given countries in Python (same equations)
"""
import os, sys
import numpy as np, pandas as pd, openpyxl
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))           # MacFiscEnergyScenarioProto/
MFM = os.path.join(ROOT, "MacroFiscalModels")
sys.path.insert(0, os.path.join(MFM, "ExcelDIGNAD"))
import build_dignad_excel as dr                                 # noqa: E402  (the DIGNAD-R model)
import outlook                                                  # noqa: E402  (levers, carbon pricing)
import nmim_live                                                # noqa: E402  (green investment, live NMIM)
import xlcharts                                                  # noqa: E402

Y0, Y1, T = dr.Y0, dr.Y1, dr.T
YEARS = list(range(Y0, Y1 + 1))
F = "Arial"; INK = "0E1B2A"; ACC = "C05621"; MUT = "5A6B7B"
IND, DETAIL, NOTMASTER = "Indicators (not master)", "Detail ->>", "Not master ->>"
HDR = PatternFill("solid", fgColor=INK); YEL = PatternFill("solid", fgColor="FFF2CC")
SEC = PatternFill("solid", fgColor="EAE3F2"); AUG = PatternFill("solid", fgColor="E3EAF2")
thin = Side(style="thin", color="C8D0D6"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(size=10, bold=False, color=INK, italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)


def load(countries=None):
    """DIGNAD-R inputs for every country of the integrated run that the dashboard has."""
    ir = pd.read_csv(os.path.join(MFM, "results", "integrated_results.csv")).set_index("iso3")
    panel = pd.read_csv(dr.PANEL).set_index("iso3")
    isos = [i for i in ir.index if i in panel.index and panel.income.get(i) in dr.INCOME]
    dr.COUNTRIES = countries or isos
    wide, years, long = dr.load_inputs()
    return wide, long, ir


def add(path, panel_long):
    wide, long, ir = load()
    names = panel_long.drop_duplicates("iso3").set_index("iso3").country   # the Dashboard's names
    wide["country"] = wide.iso3.map(names).fillna(wide.country)
    pa = pd.read_csv(os.path.join(MFM, "results", "integrated_paths.csv")).set_index(["iso3", "series"])
    pfi = outlook.pricing_inputs(wide.iso3, YEARS)
    pcodes = [c for c, *_ in outlook.PF_PAR] + [c for c, _ in outlook.PF_REF]

    wb = openpyxl.load_workbook(path)
    wb["Dashboard"].title = "Indicators"                    # the new Dashboard is the summary (outlook.py)
    LASTD = wb["Data_wide"].max_row
    rows = nmim_live.copy_engine(wb)                        # GI_Inputs, GI_Model, GI_Data

    def name(n, ref):
        wb.defined_names[n] = DefinedName(n, attr_text=ref)

    # ---- Macro_inputs: one row per country ------------------------------------------------
    wi = wb.create_sheet("Macro_inputs")
    dcodes = [c for c, *_ in dr.DASH]; acodes = [c for c, *_ in dr.AUGD]
    paths = [("cap", "pub_cap", "Budget need, capital (% GDP)"),
             ("spend", "pub_spend", "Budget need, spending (% GDP)"),
             ("rev", "rev", "Pricing revenue (% GDP)"),
             ("pyg", "gdp_dev_pct", "Integrated run: GDP vs today, programme (%)"),
             ("pygb", "gdp_dev_base_pct", "Integrated run: GDP vs today, baseline (%)"),
             ("pyd", "debt_pct_gdp", "Integrated run: debt, programme (% GDP)"),
             ("pydb", "debt_base_pct_gdp", "Integrated run: debt, baseline (% GDP)"),
             ("alt", "alt_gdp_dev_pct", "Other model: GDP vs today, programme (%)"),
             ("altd", "alt_debt_pct_gdp", "Other model: debt, programme (% GDP)")]
    DSFK = ["PV public debt / GDP", "PV ext debt / GDP", "PV ext debt / exports",
            "Ext debt service / exports", "Ext debt service / revenue"]
    DSFY = list(range(2025, 2035))
    cols = (["iso3", "country", "income", "model", "model_reason", "status", "alt_model",
             "dsf_capacity", "dsf_ci", "dsf_rating_base", "dsf_rating_programme"] + dcodes + acodes +
            [f"{p}_{y}" for p, *_ in paths for y in YEARS] +
            [f"dsf{x}{j}_{y}" for j in range(len(DSFK)) for x in "bpt" for y in DSFY] + pcodes)
    for j, c in enumerate(cols, 1):
        x = wi.cell(1, j, c); x.font = font(8, True, "FFFFFF"); x.fill = HDR
    lv = long.set_index(["iso3", "indicator", "year"]).value
    for i, row in enumerate(wide.itertuples(index=False), 2):
        rd = row._asdict(); iso = rd["iso3"]
        vals = [iso, rd["country"], rd["income"], ir.model.get(iso), ir.model_reason.get(iso), ir.status.get(iso),
                ir.alt_model.get(iso)] + [ir[c].get(iso) if c in ir else None for c in
                                          ("dsf_capacity", "dsf_ci", "dsf_rating_base", "dsf_rating_programme")]
        vals += [rd[c] for c in dcodes + acodes]
        for p, src, _ in paths:
            for y in YEARS:
                if p in ("cap", "spend", "rev"):                    # as DIGNAD-R: missing year = 0
                    v = 0.0 if y == Y0 else float(lv.get((iso, src, y), 0.0))
                elif y == Y0:                                       # integrated run starts from 2024
                    v = 0.0 if p in ("pyg", "pygb", "alt") else ir.debt_2024_pct_gdp.get(iso)
                else:
                    v = pa[str(y)].get((iso, src)) if (iso, src) in pa.index else None
                vals.append(v)
        for j_, k_ in enumerate(DSFK):
            for x, ser in zip("bpt", ("dsf_base", "dsf_prog", "dsf_thr")):
                key = (iso, f"{ser}|{k_}")
                vals += [pa[str(y)].get(key) if key in pa.index else None for y in DSFY]
        vals += [pfi.at[iso, c] for c in pcodes]
        for j, v in enumerate(vals, 1):
            v = None if (isinstance(v, float) and np.isnan(v)) else v
            wi.cell(i, j, v).font = font(8)
    NI = len(wide) + 1
    IC = {c: get_column_letter(j) for j, c in enumerate(cols, 1)}
    wi.freeze_panes = "C2"; wi.sheet_properties.tabColor = "BFBFBF"
    rng = lambda c: f"Macro_inputs!${IC[c]}$2:${IC[c]}${NI}"

    # ---- Macro_assumptions ----------------------------------------------------------------
    wa = wb.create_sheet("Macro_assumptions")
    for j, (h, w) in enumerate([("Name", 18), ("Value", 10), ("Description", 72), ("Source", 60)], 1):
        x = wa.cell(1, j, h); x.font = font(9, True, "FFFFFF"); x.fill = HDR
        wa.column_dimensions[get_column_letter(j)].width = w
    # p_pricing is replaced by the carbon pricing levers, p_programme moves to the Dashboard
    prm = [x for x in dr.PARAMS if x[0] not in ("p_pricing", "p_programme")] + outlook.PF_PRIORS
    for i, (n, v, lab, src) in enumerate(prm, 2):
        wa.cell(i, 1, n).font = font(9, True)
        c = wa.cell(i, 2, v); c.fill = YEL; c.font = font(10, True, "0000FF"); c.border = BOX
        wa.cell(i, 3, lab).font = font(9); wa.cell(i, 4, src).font = font(9, color=MUT)
        name(n, f"Macro_assumptions!$B${i}")
    wa.freeze_panes = "A2"

    # ---- Macro_model: inputs and calibration for the selected country -------------------
    wm = wb.create_sheet("Macro_model")
    wm.column_dimensions["A"].width = 30; wm.column_dimensions["B"].width = 12
    wm["A1"] = "DIGNAD-R for the country picked on the Dashboard (recursive DIGNAD-lite / DIGNAR-lite)"
    wm.sheet_properties.tabColor = "BFBFBF"
    wm["A1"].font = font(11, True)
    r = 3
    def cell(label, nm, formula, fmt="0.0000"):
        nonlocal r
        wm.cell(r, 1, label).font = font(9)
        x = wm.cell(r, 2, formula); x.font = font(9); x.number_format = fmt
        name(nm, f"Macro_model!$B${r}")
        r += 1
    cell("Row of the country in Macro_inputs", "m_row",
         f"=MATCH(sel_country,{rng('country')},0)", "0")
    # INDEX on an empty cell returns 0, not blank: test for blank first (the dashboard's rule)
    for c in dcodes + acodes + pcodes:
        ix = f"INDEX({rng(c)},m_row)"
        cell(f"input: {c}", f"i_{c}", f'=IF({ix}="","",{ix})')
    I = lambda c: f"i_{c}"
    cal = [
        ("k", "income group (1 LIC ... 4 HIC)",
         f'=MATCH(INDEX({rng("income")},m_row),{{"Low income","Lower middle income","Upper middle income","High income"}},0)'),
        ("s", "public investment efficiency",
         f'=1-IF(ISNUMBER({I("pefa_pi11_pim_score")}),0.4-({I("pefa_pi11_pim_score")}-1)/3*0.27,CHOOSE(c_k,0.4,0.27,0.27,0.13))'),
        ("Rg", "return on public capital", f"={I('return_public_capital_pct')}/100"),
        ("ig", "public investment, start", f"={I('public_investment_pct_gdp')}/100"),
        ("tax", "tax revenue", f"={I('tax_revenue_pct_gdp')}/100"),
        ("grants0", "grants, start", f"={I('grants_pct_gdp')}/100"),
        ("conc", "concessional share of debt",
         f'=IF(ISNUMBER({I("ids_multilateral_share_pct")}),{I("ids_multilateral_share_pct")}/100,CHOOSE(c_k,0.7,0.45,0.15,0))'),
        ("rstar", "real commercial rate",
         f'=IF(ISNUMBER({I("ids_private_creditor_rate_pct")}),MIN(MAX({I("ids_private_creditor_rate_pct")}/100-0.03,0.03),0.08),IF(c_k<3,0.04,0.03))'),
        ("rr", "fossil rents, share of GDP",
         f'=IF(ISNUMBER({I("fossil_rents_total_pct_gdp")}),{I("fossil_rents_total_pct_gdp")}/100,0)'),
        ("res", "resource block on (1 = DIGNAR)", "=IF(c_rr>=p_res_threshold,1,0)"),
        ("rents0", "rents, share of non-resource GDP", "=IF(c_res=1,c_rr/(1-c_rr),0)"),
        ("K0", "private capital, start", "=p_alpha*(1-c_tax)/(c_rstar+p_delta)"),
        ("KG0", "public capital, start", "=c_s*c_ig/p_delta_g"),
        ("psi", "output elasticity of public capital", "=c_Rg*c_KG0"),
        ("A", "productivity (Y 2024 = 1)", "=1/(c_KG0^c_psi*c_K0^p_alpha)"),
        ("GDP0", "GDP, start", "=1+c_rents0"),
        ("D0", "public debt, start", f"={I('gross_debt_pct_gdp')}/100*c_GDP0"),
        ("Dk0", "concessional debt, start", "=c_conc*c_D0"),
        ("int0", "interest, start", "=p_rk*c_Dk0+c_rstar*(c_D0-c_Dk0)"),
        ("rev0", "revenue, start", "=c_tax+p_gov_take*c_rents0+c_grants0"),
        ("g", "government consumption", "=MAX(0.05,c_rev0-c_ig-c_int0-p_tr_target)"),
        ("trbar", "transfers, start (balances 2024)", "=c_rev0-c_g-c_ig-c_int0"),
        ("lam_f", "fiscal rule speed",
         f"=IF(ISNUMBER({I('fiscal_rule_speed')}),{I('fiscal_rule_speed')},IF(c_res=1,p_lambda_f_res,p_lambda_f))"),
        ("ratio0", "debt ratio, start", "=c_D0/c_GDP0"),
    ]
    r += 1
    wm.cell(r, 1, "Calibration (as DIGNAD-R)").font = font(10, True); r += 1
    for k, lab, f in cal:
        cell(f"{k}: {lab}", f"c_{k}", f)

    # ---- Macro_model: the recursion -----------------------------------------------------
    r += 2
    H0 = r
    BASE = ["year", "t", "dis", "tfp", "rents", "cap", "spend", "dtax", "ramp", "cpT", "cpF", "cpE", "gdpusd", "co2", "crev"]
    LEV = ["newtax", "crev", "rebate", "green", "newpi", "newoth"]          # outlook.recursion_extra
    SV = (["KG", "K", "Y", "Kstar", "GDP", "P", "grants", "revenue", "tr", "prim", "interest", "D", "Dk", "ratio", "rc",
           "chtm"] + LEV)
    DEV = ["gdp_dev_prog", "gdp_vs_2024_P", "gdp_vs_2024_B", "debt_diff_pp", "chtm_dev_prog", "tr_diff"]
    CH = ["ch_gdp_P", "ch_gdp_B", "ch_gdp_py", "ch_gdp_pyb", "ch_gdp_alt", "ch_debt_P", "ch_debt_B", "ch_debt_py",
          "ch_debt_pyb", "ch_debt_alt",
          "ch_tr_P", "ch_tr_B", "ch_prog", "ch_rev", "ch_rents",
          "o_rev_P", "o_exp_P", "o_rev_B", "o_exp_B", "o_bal_P", "o_bal_B", "o_newtax", "o_crev", "o_grn", "o_newsp",
          "ch_gdpdev_P", "ch_gdpdev_py", "gi_dinv", "gi_bud", "ch_cpT", "ch_cpF"]
    CHL = {"ch_gdpdev_P": "Simple model, policy", "ch_gdpdev_py": "Integrated run, programme",
           "gi_dinv": "Incremental investment need (NMIM)", "gi_bud": "Budget need (NMIM)",
           "ch_cpT": "Carbon tax", "ch_cpF": "Feebate","ch_gdp_P": "Simple model, policy", "ch_gdp_B": "Simple model, baseline",
           "ch_gdp_py": "Integrated run, programme", "ch_gdp_pyb": "Integrated run, baseline",
           "ch_gdp_alt": "Other model, programme", "ch_debt_alt": "Other model, programme",
           "ch_debt_P": "Simple model, policy", "ch_debt_B": "Simple model, baseline",
           "ch_debt_py": "Integrated run, programme", "ch_debt_pyb": "Integrated run, baseline",
           "ch_tr_P": "Transfers, policy", "ch_tr_B": "Transfers, baseline",
           "ch_prog": "Programme spending", "ch_rev": "Carbon tax revenue", "ch_rents": "Fossil rents",
           "o_rev_P": "Revenue, policy", "o_exp_P": "Expenditure, policy", "o_rev_B": "Revenue, baseline",
           "o_exp_B": "Expenditure, baseline", "o_bal_P": "Balance, policy", "o_bal_B": "Balance, baseline",
           "o_newtax": "New taxes", "o_crev": "Carbon tax revenue", "o_grn": "Green public investment",
           "o_newsp": "New public investment + other spending"}
    MC = BASE + [f"B_{v}" for v in SV] + [f"P_{v}" for v in SV] + DEV + CH
    M = {c: get_column_letter(j + 2) for j, c in enumerate(MC)}         # model starts in column B
    for c in MC:
        x = wm[f"{M[c]}{H0}"]; x.value = CHL.get(c, c); x.font = font(8, True, "FFFFFF"); x.fill = HDR
        x.alignment = Alignment(wrap_text=True)
    wm.cell(H0, 1, "Model, 2024-2050 (B = baseline, P = policy: programme + levers)").font = font(9, True)
    R0 = H0 + 1
    py = lambda p, rr: f'INDEX(Macro_inputs!${IC[p + "_" + str(Y0)]}$2:${IC[p + "_" + str(Y1)]}${NI},m_row,$C{rr}+1)'
    for t in range(T):
        rr, p = R0 + t, R0 + t - 1
        v = {"year": Y0 + t, "t": t,
             "dis": f"=IF(AND(p_dis_year>0,$B{rr}=p_dis_year),1,0)",
             "tfp": f"=IF(AND(p_dis_year>0,$B{rr}>=p_dis_year),1-p_eps_a*p_rho_a^($B{rr}-p_dis_year),1)",
             "rents": f"=c_rents0*(1-p_rent_decline*MIN(1,$C{rr}/{Y1 - Y0}))",
             "cap": "=0", "spend": "=0"}
        if t > 0:                       # the climate programme: NMIM's incremental budget need, live (GI_Model)
            v["cap"], v["spend"] = (f"={f}" for f in nmim_live.budget_paths(rows, Y0 + t))
        v.update(outlook.shared_columns(rows, M, rr, t))
        for X, on, L in (("B", "0", "0"), ("P", "p_programme", "1")):
            c = lambda var: f"{M[X + '_' + var]}{rr}"
            pr = lambda var: f"{M[X + '_' + var]}{p}"
            if t == 0:
                v.update({f"{X}_KG": "=c_KG0", f"{X}_K": "=c_K0", f"{X}_Y": "=1", f"{X}_Kstar": "=c_K0",
                          f"{X}_GDP": "=c_GDP0", f"{X}_P": "=0", f"{X}_grants": "=c_grants0",
                          f"{X}_revenue": "=c_rev0", f"{X}_tr": "=c_trbar", f"{X}_prim": "=-c_int0",
                          f"{X}_interest": "=c_int0", f"{X}_D": "=c_D0", f"{X}_Dk": "=c_Dk0",
                          f"{X}_ratio": f"={c('D')}/{c('GDP')}", f"{X}_rc": "=c_rstar",
                          f"{X}_chtm": f"=(1-c_tax)*(1-p_alpha)*{c('Y')}+{c('tr')}"})
                v.update({f"{X}_{k}": "=0" for k in LEV})
                continue
            KGd = f"{pr('KG')}*(1-p_eps_g*${M['dis']}{rr})"; Kd = f"{pr('K')}*(1-p_eps_k*${M['dis']}{rr})"
            v.update({
                f"{X}_Y": f"=c_A*${M['tfp']}{rr}*({KGd})^c_psi*({Kd})^p_alpha",
                f"{X}_Kstar": f"=p_alpha*((1-c_tax)*{c('Y')}-f_tax_cap*({c('newtax')}+{c('crev')}))/(c_rstar+p_delta)",
                f"{X}_K": f"={Kd}+p_lambda_k*({c('Kstar')}-{Kd})",
                f"{X}_P": f"={on}*(${M['cap']}{rr}+${M['spend']}{rr})",
                f"{X}_KG": f"=(1-p_delta_g)*{KGd}+c_s*(c_ig+{on}*${M['cap']}{rr}+{c('newpi')}+{c('green')})",
                f"{X}_GDP": f"={c('Y')}+${M['rents']}{rr}",
                f"{X}_grants": f"=c_grants0+p_sh_grant*{c('P')}",
                f"{X}_revenue": f"=c_tax*{c('Y')}+{c('newtax')}+{c('crev')}+p_gov_take*${M['rents']}{rr}+{c('grants')}",
                f"{X}_tr": f"=c_trbar-f_rule*c_lam_f*({pr('D')}/{pr('GDP')}-c_ratio0)",
                f"{X}_prim": (f"=c_g+c_ig+{c('P')}+{c('newpi')}+{c('newoth')}+{c('green')}+{c('rebate')}"
                              f"+{c('tr')}-{c('revenue')}"),
                f"{X}_interest": f"=p_rk*{pr('Dk')}+{pr('rc')}*({pr('D')}-{pr('Dk')})",
                f"{X}_D": f"={pr('D')}+{c('interest')}+{c('prim')}",
                f"{X}_Dk": f"=c_Dk0+({pr('Dk')}-c_Dk0)*(1-p_amort)+p_sh_conc*{c('P')}",
                f"{X}_ratio": f"={c('D')}/{c('GDP')}",
                f"{X}_rc": f"=c_rstar+p_eta*MAX(0,{c('ratio')}-c_ratio0)",
                f"{X}_chtm": (f"=(1-c_tax)*(1-p_alpha)*{c('Y')}+{c('tr')}+{c('rebate')}"
                              f"-(1-p_alpha)*({c('newtax')}+{c('crev')})")})
            v.update({f"{X}_{k}": f for k, f in outlook.recursion_extra(c, pr, L, M, rr).items()})
        m = lambda col, row=rr: f"{M[col]}{row}"
        v.update({
            "gdp_dev_prog": f"=({m('P_GDP')}/{m('B_GDP')}-1)*100",
            "gdp_vs_2024_P": f"=({m('P_GDP')}/{m('P_GDP', R0)}-1)*100",
            "gdp_vs_2024_B": f"=({m('B_GDP')}/{m('B_GDP', R0)}-1)*100",
            "debt_diff_pp": f"=({m('P_ratio')}-{m('B_ratio')})*100",
            "chtm_dev_prog": f"=({m('P_chtm')}/{m('B_chtm')}-1)*100",
            "tr_diff": f"=({m('P_tr')}-{m('B_tr')})*100",
            # chart series: GDP as 2024 = 100; the integrated run's 'today's path' starts at 100 too
            "ch_gdp_P": f"={m('P_GDP')}/{m('P_GDP', R0)}*100", "ch_gdp_B": f"={m('B_GDP')}/{m('B_GDP', R0)}*100",
            "ch_gdp_py": f'=IFERROR(IF({py("pyg", rr)}="",NA(),100+{py("pyg", rr)}),NA())',
            "ch_gdp_pyb": f'=IFERROR(IF({py("pygb", rr)}="",NA(),100+{py("pygb", rr)}),NA())',
            "ch_gdp_alt": f'=IFERROR(IF({py("alt", rr)}="",NA(),100+{py("alt", rr)}),NA())',
            "ch_debt_alt": f'=IFERROR(IF({py("altd", rr)}="",NA(),{py("altd", rr)}),NA())',
            "ch_debt_P": f"={m('P_ratio')}*100", "ch_debt_B": f"={m('B_ratio')}*100",
            "ch_debt_py": f'=IFERROR(IF({py("pyd", rr)}="",NA(),{py("pyd", rr)}),NA())',
            "ch_debt_pyb": f'=IFERROR(IF({py("pydb", rr)}="",NA(),{py("pydb", rr)}),NA())',
            "ch_tr_P": f"={m('P_tr')}*100", "ch_tr_B": f"={m('B_tr')}*100",
            "ch_prog": f"={m('P_P')}*100", "ch_rev": f"={m('P_crev')}/{m('P_GDP')}*100",
            "ch_rents": f"={m('rents')}*100"})
        for X in "PB":
            exp = (f"c_g+c_ig+{m(X + '_P')}+{m(X + '_newpi')}+{m(X + '_newoth')}+{m(X + '_green')}+{m(X + '_rebate')}"
                   f"+{m(X + '_tr')}+{m(X + '_interest')}")
            v.update({f"o_rev_{X}": f"={m(X + '_revenue')}/{m(X + '_GDP')}*100",
                      f"o_exp_{X}": f"=({exp})/{m(X + '_GDP')}*100",
                      f"o_bal_{X}": f"=-({m(X + '_prim')}+{m(X + '_interest')})/{m(X + '_GDP')}*100"})
        v.update({"o_newtax": f"={m('P_newtax')}/{m('P_GDP')}*100", "o_crev": f"={m('P_crev')}/{m('P_GDP')}*100",
                  "o_grn": f"=(p_programme*{m('cap')}+{m('P_green')})/{m('P_GDP')}*100",
                  "o_newsp": f"=({m('P_newpi')}+{m('P_newoth')})/{m('P_GDP')}*100",
                  "ch_gdpdev_P": f"={m('gdp_dev_prog')}",
                  "ch_gdpdev_py": (f'=IFERROR(IF({py("pyg", rr)}="",NA(),((1+{py("pyg", rr)}/100)/(1+{py("pygb", rr)}/100)-1)*100),NA())'
                                   if t else "=0"),
                  "gi_dinv": (f"=IFERROR({nmim_live.ref(rows, 'mit.dinv.tot', Y0 + t)}/{nmim_live.ref(rows, 'mit.gdp.lvl', Y0 + t)}*100,NA())"),
                  "gi_bud": (f"=IFERROR({nmim_live.ref(rows, 'mit.dtyp.bud', Y0 + t)}/{nmim_live.ref(rows, 'mit.gdp.lvl', Y0 + t)}*100,NA())"),
                  "ch_cpT": f"={m('cpT')}", "ch_cpF": f"={m('cpF')}"})
        for col in MC:
            x = wm[f"{M[col]}{rr}"]; x.value = v[col]; x.font = font(8)
            if col not in ("year", "t", "dis"):
                x.number_format = "0.0000"
    RL = R0 + T - 1
    wm.freeze_panes = wm[f"C{R0}"]
    wm.sheet_properties.tabColor = "BFBFBF"
    at = lambda col, yr: f"Macro_model!${M[col]}${R0 + yr - Y0}"

    # ---- Macro & debt: the view ----------------------------------------------------------
    wv = wb.create_sheet("Macro & debt")
    wv.sheet_view.showGridLines = False
    for col, w in zip("ABCDEFGH", [2, 52, 11, 11, 3, 11, 11, 2]):
        wv.column_dimensions[col].width = w
    wv["B2"] = "Macro and debt: growth, public debt and debt sustainability"; wv["B2"].font = font(16, True)
    wv["B3"] = ("Two macro models side by side. The integrated run is pre-computed in Python (MacroFiscalModels): GIMF-lite, "
                "DIGNAD-lite or DIGNAR-lite by country, forward-looking, with the climate programme and PriceFeasibility's "
                "default carbon pricing. The simple model is DIGNAD-R, the same DIGNAD-lite / DIGNAR-lite equations made "
                "recursive and written in formulas: it runs the Dashboard's scenario (levers, programme, carbon pricing) and "
                "recalculates. The debt-sustainability block is the integrated run's LIC-DSF-lite. Illustrative calibration.")
    wv["B3"].font = font(9, italic=True, color=MUT); wv["B3"].alignment = Alignment(wrap_text=True, vertical="top")
    wv.merge_cells("B3:G3"); wv.row_dimensions[3].height = 66
    wv["B5"] = "Country (pick it on the Dashboard)"; wv["B5"].font = font(10, True)
    wv["C5"] = "=sel_country"; wv["C5"].font = font(12, True, ACC); wv.merge_cells("C5:G5")
    wv["B6"] = ('=IF(ISNA(m_row),"Not covered: no NMIM / integrated-run data for this country.",'
                '"Integrated run used "&INDEX(' + rng("model") + ',m_row)&" ("&INDEX(' + rng("model_reason") +
                ',m_row)&"); status: "&INDEX(' + rng("status") + ',m_row)&". Cross-check: "&INDEX(' + rng("alt_model") +
                ',m_row)&".")')
    wv["B6"].font = font(9, True, MUT); wv.merge_cells("B6:G6")
    wv["B6"].alignment = Alignment(wrap_text=True, vertical="top"); wv.row_dimensions[6].height = 26
    r = 8
    for c, txt in zip("BCDFG", ["Growth and debt", "Integrated run", "", "Simple model (live)", ""]):
        x = wv[f"{c}{r}"]; x.value = txt; x.font = font(10, True)
    for c in "BCDEFG":
        wv[f"{c}{r}"].fill = SEC
    r += 1
    for c, txt in zip("CDFG", ["2030", "2050", "2030", "2050"]):
        x = wv[f"{c}{r}"]; x.value = txt; x.font = font(9, True, MUT); x.alignment = Alignment(horizontal="center")
    r += 1
    pyv = lambda p, yr: f'INDEX({rng(p + "_" + str(yr))},m_row)'
    rows_ = [
        ("GDP vs 2024 / today's path, policy or programme (%)", lambda yr: f'=IFERROR({pyv("pyg", yr)}+0,"")',
         lambda yr: f"={at('gdp_vs_2024_P', yr)}", "0.00"),
        ("GDP, policy or programme vs baseline (%)",
         lambda yr: f'=IFERROR(((1+{pyv("pyg", yr)}/100)/(1+{pyv("pygb", yr)}/100)-1)*100,"")',
         lambda yr: f"={at('gdp_dev_prog', yr)}", "0.00"),
        ("Public debt, policy or programme (% GDP)", lambda yr: f'=IFERROR({pyv("pyd", yr)}+0,"")',
         lambda yr: f"={at('P_ratio', yr)}*100", "0.0"),
        ("Public debt, baseline (% GDP)", lambda yr: f'=IFERROR({pyv("pydb", yr)}+0,"")',
         lambda yr: f"={at('B_ratio', yr)}*100", "0.0"),
        ("Debt added (pp of GDP)", lambda yr: f'=IFERROR({pyv("pyd", yr)}-{pyv("pydb", yr)},"")',
         lambda yr: f"={at('debt_diff_pp', yr)}", "0.0"),
        ("Cross-check, other model family: GDP vs today's path (%)", lambda yr: f'=IFERROR({pyv("alt", yr)}+0,"")', None, "0.00"),
        ("Cross-check, other model family: public debt (% GDP)", lambda yr: f'=IFERROR({pyv("altd", yr)}+0,"")', None, "0.0"),
        ("Transfers, policy vs baseline: the fiscal rule (pp of GDP)", None, lambda yr: f"={at('tr_diff', yr)}", "0.00"),
        ("Hand-to-mouth consumption, policy vs baseline (%)", None, lambda yr: f"={at('chtm_dev_prog', yr)}", "0.00"),
        ("Fossil rents, resource block (% GDP)", None, lambda yr: f"={at('rents', yr)}*100", "0.0"),
    ]
    for lab, fpy, fsm, fmt in rows_:
        wv[f"B{r}"] = lab; wv[f"B{r}"].font = font(9); wv[f"B{r}"].border = BOX
        for cc, yr in (("C", 2030), ("D", 2050)):
            x = wv[f"{cc}{r}"]; x.value = fpy(yr) if fpy else "-"; x.number_format = fmt
            x.font = font(10, True); x.border = BOX; x.alignment = Alignment(horizontal="center")
        for cc, yr in (("F", 2030), ("G", 2050)):
            x = wv[f"{cc}{r}"]; x.value = f'=IF(ISNA(m_row),"",{fsm(yr)[1:]})' if fsm else "-"; x.number_format = fmt
            x.font = font(10, True); x.border = BOX; x.alignment = Alignment(horizontal="center")
        r += 1

    # debt sustainability (LIC-DSF-lite, integrated run)
    r += 1
    for c in "BCDEFG":
        wv[f"{c}{r}"].fill = SEC
    wv[f"B{r}"] = "Debt sustainability: LIC-DSF-lite (integrated run, low and lower-middle income)"; wv[f"B{r}"].font = font(10, True)
    r += 1
    gi = lambda c: f"INDEX({rng(c)},m_row)"
    wv[f"B{r}"] = (f'=IF(ISNA(m_row),"Not covered.",IF({gi("dsf_rating_base")}="","Not assessed: LIC-DSF-lite covers low '
                   f'and lower-middle income countries with a macro result.","Capacity: "&{gi("dsf_capacity")}&IF({gi("dsf_ci")}="",'
                   f'" (no CPIA: medium assumed)"," (CI "&TEXT({gi("dsf_ci")},"0.00")&")")&".   Risk rating: "&{gi("dsf_rating_base")}'
                   f'&" before the programme, "&{gi("dsf_rating_programme")}&" with it."))')
    wv[f"B{r}"].font = font(10, True, ACC); wv.merge_cells(f"B{r}:G{r}"); r += 1
    for c, t_ in zip("BCDFG", ["Indicator (max 2025-2034)", "Threshold", "Baseline", "Programme", "Breach?"]):
        x = wv[f"{c}{r}"]; x.value = t_; x.font = font(9, True, "FFFFFF"); x.fill = HDR
        x.alignment = Alignment(horizontal="center", wrap_text=True)
    wv[f"E{r}"].fill = HDR
    r += 1
    dsr = lambda x, j: (f"INDEX(Macro_inputs!${IC[f'dsf{x}{j}_{DSFY[0]}']}$2:${IC[f'dsf{x}{j}_{DSFY[-1]}']}${NI},m_row,0)")
    first = lambda x, j: f'INDEX({rng(f"dsf{x}{j}_{DSFY[0]}")},m_row)'
    for j, k_ in enumerate(DSFK):
        cells = {"B": k_, "C": f'=IFERROR(IF({first("t", j)}="","",{first("t", j)}),"")',
                 "D": f'=IFERROR(IF(COUNT({dsr("b", j)})=0,"",MAX({dsr("b", j)})),"")',
                 "F": f'=IFERROR(IF(COUNT({dsr("p", j)})=0,"",MAX({dsr("p", j)})),"")',
                 "G": f'=IF(OR(C{r}="",F{r}=""),"",IF(F{r}>C{r},"breach",IF(D{r}>C{r},"breach (also before)","-")))'}
        for c, v in cells.items():
            x = wv[f"{c}{r}"]; x.value = v; x.font = font(9, c == "G"); x.border = BOX
            x.number_format = "0.0"; x.alignment = Alignment(horizontal="left" if c == "B" else "center")
        r += 1
    r += 1
    notes = [
        "How the two macro models differ. The integrated run solves the whole path at once with forward-looking savers; GIMF-lite "
        "(most upper-middle and high-income countries) also has energy savings, private investment and the exchange rate. The simple "
        "model is DIGNAD-type for every country: each year follows from the last, one fiscal instrument (transfers).",
        "DSA: ratios in %, PV = present value. A breach of the PV of public debt benchmark or of any external threshold moves the "
        "mechanical rating up. Exports are an assumption (25-30% of GDP) until the dashboard has trade ratios. Illustrative, not a DSA.",
        "Four simple-model inputs are income-group defaults until they are in the dashboard: public investment, return on public "
        "capital, hand-to-mouth share, grants (Macro_inputs). Model assumptions: Macro_assumptions.",
    ]
    for t_ in notes:
        wv[f"B{r}"] = t_; wv[f"B{r}"].font = font(8, italic=True, color=MUT)
        wv[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top"); wv.merge_cells(f"B{r}:G{r}")
        wv.row_dimensions[r].height = 40; r += 1

    # DSA chart data, years down, hidden columns AA..
    c0, h0 = 27, 8
    labels = ["PV public debt / GDP: baseline", "PV public debt / GDP: programme", "PV public debt / GDP: benchmark",
              "PV external debt / GDP: baseline", "PV external debt / GDP: programme", "PV external debt / GDP: threshold"]
    srcs = [("b", 0), ("p", 0), ("t", 0), ("b", 1), ("p", 1), ("t", 1)]
    wv.cell(h0, c0, "year")
    for j, lab in enumerate(labels):
        wv.cell(h0, c0 + 1 + j, lab.split(": ")[1].capitalize())
    for t_, y in enumerate(DSFY):
        wv.cell(h0 + 1 + t_, c0, y)
        for j, (x, kk) in enumerate(srcs):
            yy = DSFY[0] if x == "t" else y                  # thresholds are stored once, for the first year
            wv.cell(h0 + 1 + t_, c0 + 1 + j,
                    f'=IFERROR(IF(INDEX({rng(f"dsf{x}{kk}_{yy}")},m_row)="",NA(),INDEX({rng(f"dsf{x}{kk}_{yy}")},m_row)),NA())')
    for cc in range(c0, c0 + 7):
        wv.column_dimensions[get_column_letter(cc)].hidden = True
    dl = h0 + len(DSFY)
    DSA_DATA = (wv, c0, h0, dl)

    cat = (2, R0 + 1, RL)                  # 2025-2050 on a numeric year axis
    col = lambda c: MC.index(c) + 2
    S = lambda cols_: [(col(c), H0, RL, role) for c, role in cols_]
    charts = [
        ("GDP, 2024 = 100", S([("ch_gdp_P", "policy"), ("ch_gdp_B", "baseline"), ("ch_gdp_py", "run"),
                                ("ch_gdp_pyb", "run_base"), ("ch_gdp_alt", "alt")]), "0"),
        ("Public debt, % of GDP", S([("ch_debt_P", "policy"), ("ch_debt_B", "baseline"), ("ch_debt_py", "run"),
                                      ("ch_debt_pyb", "run_base"), ("ch_debt_alt", "alt")]), "0"),
        ("Simple model: transfers under the fiscal rule, % of GDP", S([("ch_tr_P", "policy"), ("ch_tr_B", "baseline")]), "0.0"),
        ("Simple model: programme, carbon revenue and rents, % of GDP",
         S([("ch_prog", "green"), ("ch_rev", "gold"), ("ch_rents", "purple")]), "0.0"),
    ]
    for k, (title, ser, fmt) in enumerate(charts):
        wv.add_chart(xlcharts.line(wm, cat, ser, title, yfmt=fmt), f"{'IQ'[k % 2]}{2 + 16 * (k // 2)}")
    dcat = (c0, h0 + 1, dl)
    for k, (title, js) in enumerate([("DSA: PV of public debt, % of GDP", [(1, "baseline"), (2, "policy"), (3, "limit")]),
                                     ("DSA: PV of external public debt, % of GDP", [(4, "baseline"), (5, "policy"), (6, "limit")])]):
        wv.add_chart(xlcharts.line(wv, dcat, [(c0 + j, h0, dl, role) for j, role in js], title, skip=1),
                     f"{'IQ'[k]}34")

    # the other sheets: Green investment, Dashboard, Fiscal outlook, Carbon pricing, Policy settings
    wg, GDC, ghdr, glast = nmim_live.add_sheet(wb, rows, LASTD)
    ctx = dict(M=M, R0=R0, RL=RL, H0=H0, MC=MC, wm=wm, rng=rng, last_dash=LASTD, dsa=DSA_DATA,
               gi=(wg, GDC, ghdr, glast), IC=IC, NI=NI)
    outlook.add_sheets(wb, ctx)
    # the Global Fiscal Indicators profile: kept for reference, not the master (that is
    # TaxMacroFiscalDashboard/), at the far right behind its own divider
    ind = wb["Indicators"]
    ind.title = IND
    ind.data_validations.dataValidation = [d for d in ind.data_validations.dataValidation if "C4" not in str(d.sqref)]
    ind["C4"] = "=sel_country"
    ind["B2"] = "Indicators: the country profile - NOT MASTER"
    ind["B7"] = ("NOT MASTER: a copy of the Global Fiscal Indicators dashboard for reference; the master is "
                 "TaxMacroFiscalDashboard/. Country picked on the Dashboard.")
    ind["B7"].font = font(9, True, "B22222")

    def divider(title, text):
        ws = wb.create_sheet(title)
        ws.sheet_view.showGridLines = False; ws.column_dimensions["B"].width = 90
        ws["B2"] = title; ws["B2"].font = font(16, True)
        ws["B4"] = text; ws["B4"].font = font(10, color=MUT); ws["B4"].alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[4].height = 60
        ws.sheet_properties.tabColor = INK
        return ws
    divider(DETAIL, "Everything behind the Dashboard: README; Fiscal outlook (revenue and spending item by item); Macro & debt "
                    "(macro models and debt sustainability); Green investment (NMIM, live); Carbon pricing (carbon tax vs feebate); "
                    "Policy settings (secondary levers); model assumptions (GI_Inputs, Macro_assumptions, yellow); engine sheets (grey).")
    divider(NOTMASTER, "Reference only, not the master: the Global Fiscal Indicators country profile and its data (Data_wide, "
                       "Scores, Spec, Sources, Data_long). The master dashboard is maintained separately in "
                       "TaxMacroFiscalDashboard/. Data_wide also supplies the Dashboard's country list.")

    order = (["Dashboard", DETAIL, "README", "Fiscal outlook", "Macro & debt", "Green investment", "Carbon pricing",
              "Policy settings", "GI_Inputs", "Macro_assumptions", "GI_Model", "GI_Data", "Macro_model", "Macro_inputs",
              NOTMASTER, IND])
    rest = [ws for ws in wb._sheets if ws.title not in order]          # Data_wide, Scores, Spec, Sources, Data_long
    wb._sheets = [wb[n] for n in order] + rest
    wb.active = 0
    for ws in wb._sheets:
        ws.sheet_view.tabSelected = ws.title == "Dashboard"
    wb["Dashboard"].sheet_properties.tabColor = ACC
    for n in ("Fiscal outlook", "Macro & debt", "Green investment", "Carbon pricing"):
        wb[n].sheet_properties.tabColor = "1F4E79"
    for n in ("Policy settings", "GI_Inputs", "Macro_assumptions"):
        wb[n].sheet_properties.tabColor = "C9A227"
    for ws in [wb[IND]] + rest:
        ws.sheet_properties.tabColor = "D9D9D9"
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    print(f"macro_live: Macro sheets for {len(wide)} countries")
    return M, R0


if __name__ == "__main__":
    if "--check" in sys.argv:
        isos = sys.argv[sys.argv.index("--check") + 1:]
        wide, long, _ = load(isos)
        res = dr.python_check(wide, long)
        for iso, sc in res.items():
            for t in (6, 26):
                b, p = sc["B"][t], sc["P"][t]
                print(iso, Y0 + t, f"gdp_vs_2024_P={100 * (p['GDP'] / sc['P'][0]['GDP'] - 1):.6f}",
                      f"gdp_dev_prog={100 * (p['GDP'] / b['GDP'] - 1):.6f}",
                      f"debt_P={100 * p['ratio']:.6f}", f"debt_B={100 * b['ratio']:.6f}")
