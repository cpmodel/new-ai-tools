"""Build PriceFeasibility_Screen.xlsx: a light, self-contained Excel version of
price_feasibility.py. Three tabs:

  Screen     inputs (limits, priors), selected-country results, tax x feebate grid
  Countries  one row per country, all metrics as live formulas
  Data       long format, same columns as the dashboard's global_data_long.csv
             (iso3, country, region, income, is_ccdr, indicator, value, year,
             source), holding only the indicators used here, plus Ember power
             generation and emissions by fuel in the same format

Usage: python3 build_workbook.py [--check]   (--check recalculates in
LibreOffice and compares with price_feasibility.py)
"""
import argparse
import os
import subprocess
import tempfile

import numpy as np
import openpyxl
import pandas as pd
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import price_feasibility as pf

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "PriceFeasibility_Screen.xlsx")
LONG = os.path.join(pf.REPO, "TaxMacroFiscalDashboard", "build", "output", "global_data_long.csv")

PANEL_CODES = ["price_electricity_residential_usd_kwh", "price_coal_power_usd_gj",
               "price_gas_power_usd_gj", "price_gas_industry_usd_gj", "power_market_structure",
               "rise_renewable_energy_score", "coal_rents_pct_gdp", "ff_subsidy_explicit_pct_gdp",
               "fuel_exports_pct_merch"]
EMBER_CODES = {"gen_coal_twh": ("Coal", "Generation (TWh)"), "gen_gas_twh": ("Gas", "Generation (TWh)"),
               "gen_oil_twh": ("Other fossil", "Generation (TWh)"),
               "gen_total_twh": ("Total generation", "Generation (TWh)"),
               "em_coal_mt": ("Coal", "Emissions (MtCO2e)"), "em_gas_mt": ("Gas", "Emissions (MtCO2e)"),
               "em_oil_mt": ("Other fossil", "Emissions (MtCO2e)")}
LONG_COLS = ["iso3", "country", "region", "income", "is_ccdr", "indicator", "value", "year", "source"]

# Dashboard style
F = "Arial"; INK = "0E1B2A"; ACC = "C05621"; MUT = "5A6B7B"
HDR = PatternFill("solid", fgColor=INK); TINT = PatternFill("solid", fgColor="F0F3F5")
INPUT = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="C8D0D6"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

# (name, label, value, note) - every input is a workbook-level name in_<name>
LIMITS = [
    ("f1", "F1 max tariff rise (%)", 10.0, "First-year rise in residential tariff"),
    ("f2", "F2 max bottom-40% burden (% of consumption)", 0.5, "Net of rebate"),
    ("f3", "F3 max coal/oil generator payment (% of revenue)", 15.0, "Unpassed tax + net feebate payment"),
    ("rebate", "Share of tax revenue rebated to bottom 40%", 0.0, "0-1"),
    ("cap", "Rate cap ($/tCO2)", 200.0, "Upper bound for tax and feebate rates"),
]
P = pf.PRIORS
PRIORS = [
    ("elas", "Electricity demand elasticity", P["elasticity"], "Long run"),
    ("kappa", "Max share of coal+oil output displaceable by 2030", P["kappa"], ""),
    ("rise_lo", "kappa multiplier at RISE score 0", P["rise_scale"][0], ""),
    ("rise_hi", "kappa multiplier at RISE score 100", P["rise_scale"][1], ""),
    ("rise_miss", "RISE score when missing", P["rise_missing"], "Median"),
    ("gas_rep", "Share of displaced output replaced by gas", P["e_replace_share_gas"], "Rest clean"),
    ("sw_floor", "Lowest switching price ($/t)", P["switch_floor"], "Non-price barriers"),
    ("sw_miss", "Switching price when no gas price ($/t)", P["switch_missing"], ""),
    ("hr_coal", "Coal heat rate (GJ/MWh)", P["hr_coal"], ""),
    ("hr_gas", "Gas heat rate (GJ/MWh)", P["hr_gas"], "CCGT"),
    ("whole_sh", "Wholesale price / residential tariff", P["wholesale_share"], ""),
    ("res_sh", "Residential share of electricity sales", P["res_share"], ""),
    ("q40", "Bottom-40% share of residential use", P["q40"], ""),
    ("phase", "Phase-in years", P["phase_in_years"], "Tax reached in equal steps"),
    ("pt3", "Pass-through: wholesale market (PSRT 3)", P["pass_through"][3], ""),
    ("pt2", "Pass-through: single buyer (PSRT 2)", P["pass_through"][2], ""),
    ("pt1", "Pass-through: vertically integrated (PSRT 1)", P["pass_through"][1], ""),
    ("pt_hi", "Pass-through: no PSRT data, high income", P["pass_through_missing"]["High income"], ""),
    ("pt_oth", "Pass-through: no PSRT data, other", P["pass_through_missing"]["other"], ""),
    ("b40_li", "Bottom-40% electricity budget share: low income", P["b40_share"]["Low income"], "Placeholder"),
    ("b40_lmi", "... lower middle income", P["b40_share"]["Lower middle income"], "Placeholder"),
    ("b40_umi", "... upper middle income", P["b40_share"]["Upper middle income"], "Placeholder"),
    ("b40_hi", "... high income", P["b40_share"]["High income"], "Placeholder"),
]
GRID = [("tstep", "Grid: tax step ($/t)", 10.0, "Rows of the grid"),
        ("fstep", "Grid: feebate step ($/t)", 20.0, "Columns of the grid")]


def data_long():
    """Long-format rows for the indicators used, panel plus Ember."""
    lg = pd.read_csv(LONG)
    lg = lg[lg.indicator.isin(PANEL_CODES)]
    meta = pd.read_csv(LONG).drop_duplicates("iso3").set_index("iso3")[
        ["country", "region", "income", "is_ccdr"]]
    em = pd.read_csv(pf.EMBER)
    em = em[em["Area type"] == "Country or economy"]
    year = int(em.loc[em["Generation (TWh)"].notna(), "Year"].max()) - 1
    em = em[(em.Year == year) & em["ISO 3 code"].isin(meta.index)]
    rows = []
    for code, (src, col) in EMBER_CODES.items():
        s = em[em["Electricity source"] == src].set_index("ISO 3 code")[col].dropna()
        for iso, v in s.items():
            rows.append({"iso3": iso, "indicator": code, "value": round(float(v), 4), "year": year,
                         "source": f"{os.path.basename(pf.EMBER)} (Ember)"})
    e = pd.DataFrame(rows).join(meta, on="iso3")
    out = pd.concat([lg, e[LONG_COLS]], ignore_index=True)[LONG_COLS]
    return out.sort_values(["iso3", "indicator"]).reset_index(drop=True)


def countries(dl):
    """Countries with power generation and a residential tariff (as the Python tool)."""
    have = dl.pivot_table(index="iso3", columns="indicator", values="value", aggfunc="first")
    ok = (have.get("gen_total_twh", 0) > 0) & have.price_electricity_residential_usd_kwh.notna()
    em = have[["em_coal_mt", "em_gas_mt", "em_oil_mt"]].fillna(0).sum(axis=1)
    return sorted(have.index[ok & (em > 0)])


def hdr(ws, r, c, text, width=None):
    x = ws.cell(r, c, text)
    x.font = Font(name=F, size=9, bold=True, color="FFFFFF"); x.fill = HDR
    x.alignment = Alignment(wrap_text=True, vertical="center"); x.border = BOX
    if width:
        ws.column_dimensions[get_column_letter(c)].width = width


def name(wb, nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def build():
    dl = data_long()
    isos = countries(dl)
    wb = openpyxl.Workbook()
    sc = wb.active; sc.title = "Screen"
    co = wb.create_sheet("Countries")
    da = wb.create_sheet("Data")

    # ---------------------------------------------------------------- Data
    for j, c in enumerate(LONG_COLS, 1):
        hdr(da, 1, j, c, [7, 22, 24, 18, 8, 34, 10, 7, 46][j - 1])
    for i, row in enumerate(dl.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.bool_, bool)):
                v = bool(v)
            elif isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            da.cell(i, j, v).font = Font(name=F, size=9)
    da.freeze_panes = "A2"
    da.auto_filter.ref = f"A1:I{len(dl) + 1}"

    # ---------------------------------------------------------------- Screen inputs
    sc["B2"] = "Carbon tax vs feebate: feasibility screen (power sector)"
    sc["B2"].font = Font(name=F, size=16, bold=True, color=INK)
    sc["B3"] = ("Most abatement each instrument gets without breaking three political limits. "
                "Yellow cells are inputs. Screening only: priors are placeholders, see notes below.")
    sc["B3"].font = Font(name=F, size=10, italic=True, color=MUT)
    sc.column_dimensions["A"].width = 2
    sc.column_dimensions["B"].width = 46
    sc.column_dimensions["C"].width = 10
    sc.column_dimensions["D"].width = 26
    sc.column_dimensions["E"].width = 3

    r = 5
    sc.cell(r, 2, "Country (ISO3)").font = Font(name=F, size=10, bold=True, color=ACC)
    c = sc.cell(r, 3, "ZAF"); c.fill = INPUT; c.border = BOX; c.font = Font(name=F, size=11, bold=True)
    name(wb, "in_iso", "Screen!$C$5")
    dv = DataValidation(type="list", formula1=f"=Countries!$A$5:$A${4 + len(isos)}", allow_blank=False)
    sc.add_data_validation(dv); dv.add("C5")

    def block(r, title, items):
        hdr(sc, r, 2, title); hdr(sc, r, 3, "Value"); hdr(sc, r, 4, "Note")
        for i, (nm, lab, val, note) in enumerate(items, r + 1):
            sc.cell(i, 2, lab).font = Font(name=F, size=9)
            x = sc.cell(i, 3, val); x.fill = INPUT; x.border = BOX; x.font = Font(name=F, size=9)
            sc.cell(i, 4, note).font = Font(name=F, size=8, italic=True, color=MUT)
            name(wb, f"in_{nm}", f"Screen!$C${i}")
        return r + len(items) + 2

    r = block(7, "Political limits", LIMITS)
    r = block(r, "Grid", GRID)
    r = block(r, "Priors (assumptions: replace with country evidence or CPAT)", PRIORS)
    notes_row = r

    # ---------------------------------------------------------------- Countries
    co["A1"] = "Countries: all metrics are live formulas on the Data tab and the Screen inputs"
    co["A1"].font = Font(name=F, size=12, bold=True, color=INK)
    co["A2"] = ("Rates in $/tCO2. E = abatement, % of baseline fossil power CO2. Hybrid: tax searched in 11 "
                "steps up to its cap; feebate fills the remaining F3 room.")
    co["A2"].font = Font(name=F, size=9, italic=True, color=MUT)

    cols = []  # (key, header, formula template using {r} and {k:...} letters, fmt, width, group)

    def add(key, head, f, fmt="0.00", w=9, grp=""):
        cols.append((key, head, f, fmt, w, grp))

    def look(code, blank=True):
        cnt = f'COUNTIFS(Data!$A:$A,$A{{r}},Data!$F:$F,"{code}")'
        sm = f'SUMIFS(Data!$G:$G,Data!$A:$A,$A{{r}},Data!$F:$F,"{code}")'
        return f'=IF({cnt}=0,"",{sm})' if blank else f"={sm}"

    add("iso", "ISO3", None, "@", 6, "Country")
    add("cty", "Country", "=INDEX(Data!$B:$B,MATCH($A{r},Data!$A:$A,0))", "@", 20, "Country")
    add("inc", "Income", "=INDEX(Data!$D:$D,MATCH($A{r},Data!$A:$A,0))", "@", 14, "Country")
    for code, h in [("gen_coal_twh", "Coal gen TWh"), ("gen_gas_twh", "Gas gen TWh"),
                    ("gen_oil_twh", "Oil gen TWh"), ("gen_total_twh", "Total gen TWh"),
                    ("em_coal_mt", "Coal CO2 Mt"), ("em_gas_mt", "Gas CO2 Mt"), ("em_oil_mt", "Oil CO2 Mt")]:
        add(code, h, look(code, blank=False), "0.0", 8, "Data")
    for code, h in [("price_electricity_residential_usd_kwh", "Res. tariff $/kWh"),
                    ("price_coal_power_usd_gj", "Coal price power $/GJ"),
                    ("price_gas_power_usd_gj", "Gas price power $/GJ"),
                    ("price_gas_industry_usd_gj", "Gas price ind. $/GJ"),
                    ("power_market_structure", "Market structure (PSRT)"),
                    ("rise_renewable_energy_score", "RISE RE score")]:
        add(code, h, look(code), "0.000", 8, "Data")
    K = {}

    def L(key):
        return K[key] + "{r}"

    # assign letters so far, then derived columns referencing them
    def letters():
        for i, col in enumerate(cols, 1):
            K[col[0]] = get_column_letter(i)
    letters()
    add("e0", "Baseline CO2 Mt", f"={L('em_coal_mt')}+{L('em_gas_mt')}+{L('em_oil_mt')}", "0.0", 9, "Parameters")
    letters()
    add("eavg", "Avg intensity t/MWh", f"={L('e0')}/{L('gen_total_twh')}", "0.000", 8, "Parameters")
    add("ecoal", "Coal t/MWh", f"=IF({L('gen_coal_twh')}>0,{L('em_coal_mt')}/{L('gen_coal_twh')},0.85)", "0.000", 7, "Parameters")
    add("egas", "Gas t/MWh", f"=IF({L('gen_gas_twh')}>0,{L('em_gas_mt')}/{L('gen_gas_twh')},0.4)", "0.000", 7, "Parameters")
    add("ehi", "Coal+oil t/MWh",
        f"=IF({L('gen_coal_twh')}+{L('gen_oil_twh')}>0,({L('em_coal_mt')}+{L('em_oil_mt')})/({L('gen_coal_twh')}+{L('gen_oil_twh')}),0)",
        "0.000", 7, "Parameters")
    add("pres", "Tariff $/MWh", f"={L('price_electricity_residential_usd_kwh')}*1000", "0.0", 7, "Parameters")
    letters()
    m = L("power_market_structure")
    add("phi", "Pass-through",
        f'=IF({m}=3,in_pt3,IF({m}=2,in_pt2,IF({m}=1,in_pt1,IF({L("inc")}="High income",in_pt_hi,in_pt_oth))))',
        "0.00", 7, "Parameters")
    rise = L("rise_renewable_energy_score")
    letters()
    add("smax", "Max substitution share",
        f"=IF({L('e0')}>0,in_kappa*(in_rise_lo+(in_rise_hi-in_rise_lo)*IF({rise}=\"\",in_rise_miss,{rise})/100)"
        f"*({L('gen_coal_twh')}+{L('gen_oil_twh')})*MAX({L('ehi')}-in_gas_rep*{L('egas')},0)/{L('e0')},0)",
        "0.000", 8, "Parameters")
    pg = f'IF({L("price_gas_power_usd_gj")}="",{L("price_gas_industry_usd_gj")},{L("price_gas_power_usd_gj")})'
    add("tsw", "Switching price $/t",
        f'=IF({pg}="",MAX(in_sw_floor,in_sw_miss),MAX(in_sw_floor,({pg}*in_hr_gas-{L("price_coal_power_usd_gj")}*in_hr_coal)'
        f'/MAX({L("ecoal")}-{L("egas")},0.1)))', "0.0", 8, "Parameters")
    inc = L("inc")
    add("b40", "B40 elec. budget share",
        f'=IF({inc}="Low income",in_b40_li,IF({inc}="Lower middle income",in_b40_lmi,'
        f'IF({inc}="Upper middle income",in_b40_umi,IF({inc}="High income",in_b40_hi,0.03))))', "0.000", 7, "Parameters")
    add("hish", "Coal+oil gen share", f"=({L('gen_coal_twh')}+{L('gen_oil_twh')})/{L('gen_total_twh')}", "0.00", 7, "Parameters")
    letters()
    pw = f"(in_whole_sh*{L('pres')})"
    add("a1", "F1 per $ tax", f"=100*{L('phi')}*{L('eavg')}/{L('pres')}/in_phase", "0.000", 7, "Coefficients")
    add("a2", "F2 per $ tax",
        f"=100*{L('b40')}*{L('phi')}*{L('eavg')}/{L('pres')}*(1-IF({L('phi')}>0,MIN(in_rebate/({L('phi')}*in_res_sh*in_q40),1),1))",
        "0.0000", 7, "Coefficients")
    add("a3t", "F3 per $ tax", f"=IF({L('hish')}>0,100*{L('ehi')}*(1-{L('phi')})/{pw},0)", "0.000", 7, "Coefficients")
    add("a3f", "F3 per $ feebate", f"=IF({L('hish')}>0,100*MAX({L('ehi')}-{L('eavg')},0)/{pw},0)", "0.000", 7, "Coefficients")
    letters()

    def E(tt, tf):
        return (f"100*(1-(1-{L('smax')}*(1-EXP(-(({tt})+({tf}))/{L('tsw')})))"
                f"*(1-MIN(ABS(in_elas)*{L('phi')}*{L('eavg')}*({tt})/{L('pres')},1)))")

    def cap(a, lim):
        return f"IF({L(a)}>0,{lim}/{L(a)},in_cap)"

    add("c1", "Tax cap from F1", f"={cap('a1', 'in_f1')}", "0.0", 7, "Tax only")
    add("c2", "Tax cap from F2", f"={cap('a2', 'in_f2')}", "0.0", 7, "Tax only")
    add("c3", "Tax cap from F3", f"={cap('a3t', 'in_f3')}", "0.0", 7, "Tax only")
    letters()
    add("tt", "Best tax $/t", f"=MIN(in_cap,{L('c1')},{L('c2')},{L('c3')})", "0.0", 7, "Tax only")
    letters()
    add("tE", "Tax: E %", "=" + E(L("tt"), "0"), "0.00", 7, "Tax only")
    add("tbind", "Tax: binding",
        f'=IF({L("tt")}>=in_cap-1E-9,"rate cap",TRIM(IF(ABS({L("tt")}-{L("c1")})<1E-9,"F1 ","")'
        f'&IF(ABS({L("tt")}-{L("c2")})<1E-9,"F2 ","")&IF(ABS({L("tt")}-{L("c3")})<1E-9,"F3","")))', "@", 8, "Tax only")
    letters()
    add("tsh", "Tax: +1pp F1 buys (pp E)",
        "=" + E(f"MIN(in_cap,IF({L('a1')}>0,(in_f1+1)/{L('a1')},in_cap),{L('c2')},{L('c3')})", "0") + f"-{L('tE')}",
        "0.00", 8, "Tax only")
    add("ft", "Best feebate $/t", f"=MIN(in_cap,{cap('a3f', 'in_f3')})", "0.0", 7, "Feebate only")
    letters()
    add("fE", "Feebate: E %", "=" + E("0", L("ft")), "0.00", 7, "Feebate only")
    add("fbind", "Feebate: binding", f'=IF({L("ft")}>=in_cap-1E-9,"rate cap","F3")', "@", 8, "Feebate only")
    add("E50t", "E at $50 tax", "=" + E("50", "0"), "0.00", 7, "At $50/t")
    add("E50f", "E at $50 feebate", "=" + E("0", "50"), "0.00", 7, "At $50/t")
    letters()
    add("s50", "Feebate/tax abatement ratio", f"=IF({L('E50t')}>0,{L('E50f')}/{L('E50t')},\"\")", "0.00", 8, "At $50/t")
    add("F1_50", "F1 at $50 tax", f"=50*{L('a1')}", "0.0", 7, "At $50/t")
    letters()
    # hybrid search: k = 0..10, tax = best tax * k/10, feebate fills remaining F3 room
    first_hyb = len(cols) + 1 + 6  # summary columns come first, helpers after
    hyb_cols = ["hE", "htt", "htf", "hF1", "hF2", "hF3", "hbind", "hmt"]
    tf_keys = [f"tf{k}" for k in range(11)]
    e_keys = [f"E{k}" for k in range(11)]
    # reserve summary columns, fill formulas after helper letters are known
    for key, head, fmt, w in [("hE", "Hybrid: E %", "0.00", 7), ("htt", "Hybrid: tax $/t", "0.0", 7),
                              ("htf", "Hybrid: feebate $/t", "0.0", 7), ("hF1", "Hybrid: F1", "0.0", 6),
                              ("hF2", "Hybrid: F2", "0.00", 6), ("hF3", "Hybrid: F3", "0.0", 6),
                              ("hbind", "Hybrid: binding", "@", 8), ("hmt", "Hybrid: Mt saved", "0.0", 8)]:
        add(key, head, "", fmt, w, "Hybrid (best mix)")
    for k in range(11):
        tt = f"{L('tt')}*{k}/10"
        add(f"tf{k}", f"Feebate at k={k}",
            f"=MAX(0,MIN(in_cap,IF({L('a3f')}>0,(in_f3-{L('a3t')}*{tt})/{L('a3f')},in_cap)))",
            "0.0", 6, "Hybrid search: feebate")
    letters()
    for k in range(11):
        add(f"E{k}", f"E at k={k}", "=" + E(f"{L('tt')}*{k}/10", L(f"tf{k}")), "0.00", 6, "Hybrid search: E")
    letters()
    erng = f"{K['E0']}{{r}}:{K['E10']}{{r}}"
    frng = f"{K['tf0']}{{r}}:{K['tf10']}{{r}}"
    kstar = f"(MATCH({L('hE')},{erng},0)-1)"
    formulas = {
        "hE": f"=MAX({erng})",
        "htt": f"={L('tt')}*{kstar}/10",
        "htf": f"=INDEX({frng},1,{kstar}+1)",
        "hF1": f"={L('a1')}*{L('htt')}",
        "hF2": f"={L('a2')}*{L('htt')}",
        "hF3": f"={L('a3t')}*{L('htt')}+{L('a3f')}*{L('htf')}",
        "hbind": (f'=TRIM(IF(AND({L("htt")}>0,ABS({L("htt")}-{L("c1")})<1E-6),"F1 ","")'
                  f'&IF(AND({L("htt")}>0,ABS({L("htt")}-{L("c2")})<1E-6),"F2 ","")'
                  f'&IF({L("hF3")}>=in_f3-1E-6,"F3 ","")&IF({L("htf")}>=in_cap-1E-9,"rate cap",""))'),
        "hmt": f"={L('e0')}*{L('hE')}/100",
    }
    cols = [(k, h, formulas.get(k, f), fm, w, g) for k, h, f, fm, w, g in cols]

    R0 = 5  # first data row
    prev = None
    for j, (key, head, f, fmt, w, grp) in enumerate(cols, 1):
        if grp != prev:
            x = co.cell(3, j, grp); x.font = Font(name=F, size=8, bold=True, color=ACC)
            prev = grp
        hdr(co, 4, j, head, w)
        for i, iso in enumerate(isos):
            r = R0 + i
            v = iso if key == "iso" else f.format(r=r)
            x = co.cell(r, j, v); x.font = Font(name=F, size=9); x.number_format = fmt
            if grp.startswith("Hybrid (") or key in ("tE", "fE"):
                x.fill = TINT
    co.row_dimensions[4].height = 42
    co.freeze_panes = co.cell(R0, 4)
    last = R0 + len(isos) - 1
    co.auto_filter.ref = f"A4:{get_column_letter(len(cols))}{last}"
    for j, col in enumerate(cols, 1):
        if col[5].startswith("Hybrid search"):
            co.column_dimensions[get_column_letter(j)].outlineLevel = 1
            co.column_dimensions[get_column_letter(j)].hidden = True

    # ---------------------------------------------------------------- Screen: selected country
    name(wb, "sel_row", "Screen!$H$5")
    sc["G5"] = "Row in Countries"; sc["G5"].font = Font(name=F, size=8, color=MUT)
    sc["H5"] = f"=MATCH(in_iso,Countries!$A$1:$A${last},0)"; sc["H5"].font = Font(name=F, size=8, color=MUT)
    sc.column_dimensions["F"].width = 3
    sc.column_dimensions["G"].width = 30

    def pick(key):
        return f"=INDEX(Countries!${K[key]}$1:${K[key]}${last},sel_row)"

    out = [("Country", "cty", "@"), ("Income", "inc", "@"), ("Baseline power CO2 (Mt)", "e0", "0.0"),
           ("Average intensity (t/MWh)", "eavg", "0.000"), ("Residential tariff ($/MWh)", "pres", "0.0"),
           ("Pass-through", "phi", "0.00"), ("Switching price ($/t)", "tsw", "0.0"),
           ("Max substitution share", "smax", "0.000"), ("Feebate/tax abatement at $50", "s50", "0.00"),
           ("Tariff rise from $50 tax (%)", "F1_50", "0.0")]
    hdr(sc, 7, 7, "Selected country"); hdr(sc, 7, 8, "")
    for i, (lab, key, fmt) in enumerate(out, 8):
        sc.cell(i, 7, lab).font = Font(name=F, size=9)
        x = sc.cell(i, 8, pick(key)); x.number_format = fmt; x.font = Font(name=F, size=9, bold=True)
    sc.column_dimensions["H"].width = 16

    r = 8 + len(out) + 1
    heads = ["Best under limits", "E (%)", "Tax $/t", "Feebate $/t", "F1", "F2", "F3", "Binding"]
    for j, h in enumerate(heads, 7):
        hdr(sc, r, j, h)
    rows = [("Carbon tax / ETS", "tE", "tt", None, "tbind"),
            ("Feebate", "fE", None, "ft", "fbind"),
            ("Hybrid (tax + feebate)", "hE", "htt", "htf", "hbind")]
    for i, (lab, e, t, fb, b) in enumerate(rows, r + 1):
        sc.cell(i, 7, lab).font = Font(name=F, size=9, bold=True)
        tt = pick(t) if t else "=0"
        ff = pick(fb) if fb else "=0"
        vals = [pick(e), tt, ff, None, None, None, pick(b)]
        for j, v in enumerate(vals, 8):
            if v is not None:
                x = sc.cell(i, j, v); x.font = Font(name=F, size=9)
                x.number_format = "0.00" if j == 8 else ("0.0" if j < 14 else "@")
        tc, fc = f"I{i}", f"J{i}"
        for j, a in [(11, "a1"), (12, "a2")]:
            x = sc.cell(i, j, f"=INDEX(Countries!${K[a]}$1:${K[a]}${last},sel_row)*{tc}")
            x.number_format = "0.00"; x.font = Font(name=F, size=9)
        x = sc.cell(i, 13, f"=INDEX(Countries!${K['a3t']}$1:${K['a3t']}${last},sel_row)*{tc}"
                           f"+INDEX(Countries!${K['a3f']}$1:${K['a3f']}${last},sel_row)*{fc}")
        x.number_format = "0.0"; x.font = Font(name=F, size=9)
    for col in "IJKLMN":
        sc.column_dimensions[col].width = 10
    r += 5
    sc.cell(r, 7, "+1pp on F1 buys, tax only (pp E)").font = Font(name=F, size=9)
    x = sc.cell(r, 8, pick("tsh")); x.number_format = "0.00"; x.font = Font(name=F, size=9, bold=True)

    # grid
    g0 = r + 2
    sc.cell(g0, 7, "Abatement E (%) by tax (rows) and feebate (columns). Green = within all three limits; "
                   "bold border = best feasible cell.").font = Font(name=F, size=9, bold=True, color=INK)
    sc.cell(g0 + 1, 7, "Tax $/t \\ Feebate $/t").font = Font(name=F, size=8, bold=True, color="FFFFFF")
    sc.cell(g0 + 1, 7).fill = HDR
    # selected-country parameters for the grid, from Countries
    par = {}
    pr = g0 + 16
    sc.cell(pr, 7, "Grid parameters (from Countries)").font = Font(name=F, size=8, italic=True, color=MUT)
    for i, key in enumerate(["smax", "tsw", "phi", "eavg", "pres", "a1", "a2", "a3t", "a3f"]):
        sc.cell(pr + 1, 8 + i, key).font = Font(name=F, size=8, color=MUT)
        x = sc.cell(pr + 2, 8 + i, pick(key)); x.font = Font(name=F, size=8, color=MUT)
        x.number_format = "0.000"
        par[key] = f"${get_column_letter(8 + i)}${pr + 2}"
    NR, NC = 11, 11
    for j in range(NC):
        x = sc.cell(g0 + 1, 8 + j, f"=in_fstep*{j}"); x.fill = HDR
        x.font = Font(name=F, size=8, bold=True, color="FFFFFF"); x.number_format = "0"
    for i in range(NR):
        x = sc.cell(g0 + 2 + i, 7, f"=in_tstep*{i}"); x.fill = HDR
        x.font = Font(name=F, size=8, bold=True, color="FFFFFF"); x.number_format = "0"
        for j in range(NC):
            tt = f"$G{g0 + 2 + i}"; tf = f"{get_column_letter(8 + j)}${g0 + 1}"
            e = (f"=100*(1-(1-{par['smax']}*(1-EXP(-({tt}+{tf})/{par['tsw']})))"
                 f"*(1-MIN(ABS(in_elas)*{par['phi']}*{par['eavg']}*{tt}/{par['pres']},1)))")
            x = sc.cell(g0 + 2 + i, 8 + j, e); x.number_format = "0.0"; x.font = Font(name=F, size=9)
            x.border = BOX; x.alignment = Alignment(horizontal="center")
    gr = f"H{g0 + 2}:{get_column_letter(7 + NC)}{g0 + 1 + NR}"
    tl_t, tl_f = f"$G{g0 + 2}", f"H${g0 + 1}"
    feas = (f"AND({par['a1']}*{tl_t}<=in_f1,{par['a2']}*{tl_t}<=in_f2,"
            f"{par['a3t']}*{tl_t}+{par['a3f']}*{tl_f}<=in_f3)")
    green = PatternFill("solid", fgColor="CDE8D2", bgColor="CDE8D2")
    grey = PatternFill("solid", fgColor="EEEEEE", bgColor="EEEEEE")
    best_border = Border(left=Side("medium", color=INK), right=Side("medium", color=INK),
                         top=Side("medium", color=INK), bottom=Side("medium", color=INK))
    # best feasible value in the grid, computed in a helper cell
    bc = f"$H${pr + 4}"
    sc.cell(pr + 4, 7, "Best feasible E in grid").font = Font(name=F, size=8, italic=True, color=MUT)
    cells = []
    for i in range(NR):
        for j in range(NC):
            t = f"$G{g0 + 2 + i}"; fcell = f"{get_column_letter(8 + j)}${g0 + 1}"
            ok = (f"AND({par['a1']}*{t}<=in_f1,{par['a2']}*{t}<=in_f2,"
                  f"{par['a3t']}*{t}+{par['a3f']}*{fcell}<=in_f3)")
            cells.append(f"IF({ok},{get_column_letter(8 + j)}{g0 + 2 + i},0)")
    x = sc.cell(pr + 4, 8, "=MAX(" + ",".join(cells) + ")")
    x.font = Font(name=F, size=8, color=MUT); x.number_format = "0.00"
    sc.conditional_formatting.add(gr, FormulaRule(formula=[f"AND({feas},ABS(H{g0 + 2}-{bc})<1E-9)"],
                                                  fill=green, border=best_border, font=Font(bold=True)))
    sc.conditional_formatting.add(gr, FormulaRule(formula=[feas], fill=green))
    sc.conditional_formatting.add(gr, FormulaRule(formula=[f"NOT({feas})"], fill=grey,
                                                  font=Font(color="9A9A9A")))

    # notes
    r = notes_row
    hdr(sc, r, 2, "Notes"); hdr(sc, r, 3, ""); hdr(sc, r, 4, "")
    notes = [
        "What: most power-sector CO2 cut from a carbon tax/ETS, a feebate (= output-based allocation, TPS) "
        "or a hybrid, without exceeding three political limits.",
        "E = 1-(1-A_sub)(1-A_dem). A_sub = s_max(1-exp(-(tax+feebate)/switching price)): cleaner generation, "
        "driven by the marginal price. A_dem = |elasticity| x tariff rise: less use, driven by the tax only.",
        "F1 = tariff rise = pass-through x avg intensity x tax / tariff. F2 = bottom-40% budget share x F1 x "
        "(1 - rebate cover); full cover needs pass-through x residential share x q40 of revenue.",
        "F3 = [tax x coal/oil intensity x (1-pass-through) + feebate x (coal/oil intensity - average)] / "
        "wholesale price: what coal and oil generators pay.",
        "Data: same long format as the dashboard's global_data_long.csv. To refresh, paste new rows on "
        "Data (same columns); add a country by adding its ISO3 in Countries col A and copying a row down.",
        "Screening only. Priors above are placeholders. One sector, one period (~2030). Python version and "
        "method: price_feasibility.py, README.md, NOTES.md.",
    ]
    for i, t in enumerate(notes, r + 1):
        sc.merge_cells(start_row=i, start_column=2, end_row=i, end_column=4)
        x = sc.cell(i, 2, t); x.font = Font(name=F, size=8, color=MUT)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        sc.row_dimensions[i].height = 34
    sc.page_setup.orientation = "landscape"
    sc.page_setup.fitToWidth, sc.page_setup.fitToHeight = 1, 1
    sc.sheet_properties.pageSetUpPr.fitToPage = True
    sc.freeze_panes = "A5"
    sc.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    return OUT, K, len(isos)


def check(path, K, n):
    tmp = tempfile.mkdtemp()
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", tmp, path],
                   check=True, capture_output=True, timeout=600)
    wb = openpyxl.load_workbook(os.path.join(tmp, os.path.basename(path)), data_only=True)
    co = wb["Countries"]
    from openpyxl.utils import column_index_from_string as ci
    rows = []
    for r in range(5, 5 + n):
        g = lambda k: co.cell(r, ci(K[k])).value
        rows.append({"iso3": g("iso"), "tax_E": g("tE"), "feebate_E": g("fE"), "hybrid_E": g("hE"),
                     "E_tax50": g("E50t"), "E_feebate50": g("E50f")})
    x = pd.DataFrame(rows).set_index("iso3")
    errs = [(r, c) for r in range(1, co.max_row + 1) for c in range(1, co.max_column + 1)
            if isinstance(co.cell(r, c).value, str) and co.cell(r, c).value.startswith("#")]
    py = pf.country_metrics(pf.build_calibration())
    j = x.join(py[["tax_E", "feebate_E", "hybrid_E", "E_tax50", "E_feebate50"]], rsuffix="_py", how="inner")
    print(f"Countries: {n} in workbook, {len(j)} matched with Python; formula errors: {len(errs)}")
    for c in ["E_tax50", "E_feebate50", "tax_E", "feebate_E", "hybrid_E"]:
        d = (j[c].astype(float) - j[c + "_py"]).abs()
        print(f"  {c:12s} max |xlsx - py| = {d.max():.3f} pp  (median {d.median():.4f})")
    sc = wb["Screen"]
    print("Screen sample:", [sc.cell(r, 8).value for r in range(8, 12)])


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    a = ap.parse_args()
    path, K, n = build()
    print(f"Wrote {path} ({os.path.getsize(path) // 1024} KB, {n} countries)")
    if a.check:
        check(path, K, n)
