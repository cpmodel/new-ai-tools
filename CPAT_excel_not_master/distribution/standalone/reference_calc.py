#!/usr/bin/env python3
"""
Pure-Python reference implementation of Steps 2-12 (pseudocode), used to
validate the Excel LAMBDA formula design against the source workbook's own
cached values *before* encoding the logic as (harder-to-debug) Excel formula
strings. Not shipped in the workbook itself.
"""
import pickle

HERE = "/home/user/CPAT-AI-Local/cpat_excel/Distribution/standalone"
with open(f"{HERE}/egypt_data.pkl", "rb") as f:
    DATA = pickle.load(f)
with open("/tmp/claude-0/-home-user-CPAT-AI-Local/117e96f5-0306-5814-b5e4-f0cac77f3abb/scratchpad/dist_grid.pkl", "rb") as f:
    GRID = pickle.load(f)

hh_header, hh_rows = DATA["hhsurvey"]
el_header, el_rows = DATA["hh_elast"]
HH_IDX = {h: i for i, h in enumerate(hh_header)}
EL_IDX = {h: i for i, h in enumerate(el_header)}


def _key_match(rows, idx, sample, qtype, stat, dnum):
    for row in rows:
        if (row[idx["sample"]] == sample and row[idx["type"]] == qtype
                and row[idx["stat_type"]] == stat and row[idx["quant_cons"]] == dnum):
            return row
    raise KeyError((sample, qtype, stat, dnum))


def bshare(colname, decile, sample, stat):
    qtype = "Basket" if decile == 0 else "Deciles"
    dnum = 9999 if decile == 0 else decile
    row = _key_match(hh_rows, HH_IDX, sample, qtype, stat, dnum)
    return row[HH_IDX[colname]]


def elast(colname, decile, sample, stat):
    qtype = "Basket" if decile == 0 else "Deciles"
    dnum = 9999 if decile == 0 else decile
    row = _key_match(el_rows, EL_IDX, sample, qtype, stat, dnum)
    return row[EL_IDX[colname]]


PRICE_DIRECT = DATA["price_change_direct"]
PRICE_INDIRECT = DATA["price_change_indirect"]
ELAST_ADJ = DATA["elasticity_adjustment"]  # item -> [decile1..decile10] factor


def elast_adj_factor(item, decile):
    # col16 doubles as both the basket factor and decile-1 factor in the source
    idx = 0 if decile == 0 else decile - 1
    return ELAST_ADJ[item][idx]

FUEL_SHARE_COL = {c: f"{c}_share" for c in
                   ["coa", "ely", "nga", "oil", "gso", "die", "ker", "lpg", "ccl", "ethanol", "fwd"]}
FUEL_ELAST_COL = {c: f"{c}_elasticity" for c in ["coa", "ely", "nga", "oil", "gso", "die", "ker", "lpg"]}
CAT_SHARE_COL = {"app": "appliances_share", "che": "chemicals_share", "clo": "clothing_share",
                  "com": "communications_share", "edu": "education_share", "food": "food_share",
                  "hea": "health_srv_share", "hou": "housing_share", "oth": "other_share",
                  "pap": "paper_share", "pha": "pharma_share", "ret": "rectourism_share",
                  "teq": "transp_eqt_share", "tpu": "transp_pub_share"}
CAT_ELAST_COL = {"app": "appliances_elasticity", "che": "chemicals_elasticity", "clo": "clothing_elasticity",
                  "com": "communications_elasticity", "edu": "education_elasticity", "food": "food_elasticity",
                  "hea": "health_srv_elasticity", "hou": "housing_elasticity", "oth": "other_elasticity",
                  "pap": "paper_elasticity", "pha": "pharma_elasticity", "ret": "rectourism_elasticity",
                  "teq": "transp_eqt_elasticity", "tpu": "transp_pub_elasticity"}

# Egypt switches (Section B, confirmed values)
ADJUST_BEHAVIOR = False
BEHAVIOR_FACTOR = 1.0
ADJUST_DWL = False
EXEMPT = False


def behavior_adj():
    return BEHAVIOR_FACTOR if ADJUST_BEHAVIOR else 1.0


def dwl(sharecol, elastcol, decile, sample, stat, price_pct):
    if not ADJUST_DWL:
        return 0.0
    e = elast(elastcol, decile, sample, stat)
    bs = bshare(sharecol, decile, sample, stat)
    return 0.5 * e * (price_pct / 100) ** 2 * (bs / 100) * 100


def direct_effect(fuel, decile, sample, stat):
    price = PRICE_DIRECT[fuel] * elast_adj_factor(fuel, decile)
    bs = bshare(FUEL_SHARE_COL[fuel], decile, sample, stat)
    d = dwl(FUEL_SHARE_COL[fuel], FUEL_ELAST_COL[fuel], decile, sample, stat, price)
    return (bs / 100) * (price / 100) * behavior_adj() * 100 - d


def indirect_effect(cat, decile, sample, stat):
    price = PRICE_INDIRECT[cat] * elast_adj_factor(cat, decile)
    bs = bshare(CAT_SHARE_COL[cat], decile, sample, stat)
    d = dwl(CAT_SHARE_COL[cat], CAT_ELAST_COL[cat], decile, sample, stat, price)
    return (bs / 100) * (price / 100) * behavior_adj() * 100 - d


SCALARS = {i["label"]: i["value"] for i in DATA["scalars"] if i["kind"] == "item"}
POP_ADJ_FACTOR = SCALARS["Population Adj. Factor"]
HH_CONS_ADJ_FACTOR = SCALARS["Household Consumption Adj. Factor"]
GDP_DEFLATOR_RATIO_SVY = SCALARS["GDP Deflator Ratio (HH Survey)"]


def adj_pop(decile, sample):
    return bshare("popw", decile, sample, "mean") * POP_ADJ_FACTOR


def adj_cons_pc(decile, sample):
    return bshare("cons_pc_acrent", decile, sample, "mean") * GDP_DEFLATOR_RATIO_SVY * HH_CONS_ADJ_FACTOR


def adj_cons_tot(decile, sample):
    return adj_cons_pc(decile, sample) * adj_pop(decile, sample)


def tax_burden_pre(decile, sample):
    return total_effect(decile, sample, "mean") / 100 * adj_cons_tot(decile, sample)


PIT_SHARE_COL = "pit_share_income"  # PITSourceSwitch = "income" for Egypt
PIT_BASELINE_LCU = SCALARS["Baseline PIT Revenues - Analysis Year"]  # first match = %GDP row; fix below
for item in DATA["scalars"]:
    if item["kind"] == "item" and item["label"] == "Baseline PIT Revenues - Analysis Year" and item["unit"] and "LCU" in item["unit"]:
        PIT_BASELINE_LCU = item["value"]
REV_LABOR_TAX_LCU = None
for item in DATA["scalars"]:
    if item["kind"] == "item" and item["label"] == "CP Revenue - Labor Tax Reductions - Analysis Year" and item["unit"] == "Real 2026 LCU":
        REV_LABOR_TAX_LCU = item["value"]


PIT_SHARE_SUM = sum(bshare(PIT_SHARE_COL, d, "Overall", "mean") for d in range(1, 11))


def pit_liability(decile):
    share = bshare(PIT_SHARE_COL, decile, "Overall", "mean") / PIT_SHARE_SUM * 100
    return share / 100 * PIT_BASELINE_LCU


def pit_liability_pc(decile):
    return pit_liability(decile) / adj_pop(decile, "Overall")


TOTAL_ADJ_POP = sum(adj_pop(d, "Overall") for d in range(1, 11))
MAX_TRANSFER_PC = REV_LABOR_TAX_LCU / TOTAL_ADJ_POP
SHORTFALL_TOTAL = sum(max(0, MAX_TRANSFER_PC - pit_liability_pc(d)) * adj_pop(d, "Overall") for d in range(1, 11))
ADDITIONAL_PC = SHORTFALL_TOTAL / TOTAL_ADJ_POP


def pit_reduction_personal_allowance(decile):
    pc = min(MAX_TRANSFER_PC, pit_liability_pc(decile)) + ADDITIONAL_PC
    return pc * adj_pop(decile, "Overall")


REV_TARGETED_TRANSFER_LCU = None
REV_PUBLIC_INVEST_LCU = None
REV_CURRENT_SPEND_LCU = None
for item in DATA["scalars"]:
    if item["kind"] != "item" or item["unit"] != "Real 2026 LCU":
        continue
    if item["label"] == "CP Revenue - Targeted Transfer - Analysis Year":
        REV_TARGETED_TRANSFER_LCU = item["value"]
    elif item["label"] == "CP Revenue - Public Investment - Analysis Year":
        REV_PUBLIC_INVEST_LCU = item["value"]
    elif item["label"] == "CP Revenue - Current Spending - Analysis Year":
        REV_CURRENT_SPEND_LCU = item["value"]

TOTAL_UNSERVED = sum(adj_pop(d, "Overall") / POP_ADJ_FACTOR * (1 - bshare("all_acs_share", d, "Overall", "mean") / 100) for d in range(1, 11))


def infra_share(decile):
    pop = adj_pop(decile, "Overall") / POP_ADJ_FACTOR
    access = bshare("all_acs_share", decile, "Overall", "mean")
    return (pop * (1 - access / 100)) / TOTAL_UNSERVED


ASPIRE_PIVOT = {}
h, rows = DATA["aspire"]
idx_code = h.index("Series_Code")
idx_val = h.index("(firstnm) Value")
for row in rows:
    code = row[idx_code]
    body = code.split(".", 2)
    program = body[1][len("per_"):]
    q = int(body[2].split("_")[1][1:])
    ASPIRE_PIVOT.setdefault(program, {})[q] = row[idx_val]


def aspire_pc(program, decile):
    q = -(-decile // 2)  # roundup(decile/2)
    return ASPIRE_PIVOT[program][q]


def aspire_share(program, decile):
    total = sum(aspire_pc(program, d) * adj_pop(d, "Overall") for d in range(1, 11))
    return (aspire_pc(program, decile) * adj_pop(decile, "Overall")) / total


def targeted_transfer(decile):
    return REV_TARGETED_TRANSFER_LCU * infra_share(decile)


def public_investment(decile):
    return REV_PUBLIC_INVEST_LCU * infra_share(decile)


def current_spending(decile):
    return REV_CURRENT_SPEND_LCU * aspire_share("allsp", decile)


def amount_recycled(decile):
    return (pit_reduction_personal_allowance(decile) + targeted_transfer(decile)
            + public_investment(decile) + current_spending(decile))


def net_effect(decile, sample):
    return (total_effect(decile, sample, "mean") / 100
            - amount_recycled(decile) / adj_cons_tot(decile, sample)) * 100


def post_cp_excl(decile, sample):
    return adj_cons_tot(decile, sample) * (1 - total_effect(decile, sample, "mean") / 100)


def post_cp_incl(decile, sample):
    return post_cp_excl(decile, sample) + amount_recycled(decile)


def gini(shares):
    """shares: list of 10 (non-cumulative) decile shares, poorest->richest."""
    cum = []
    running = 0.0
    for s in shares:
        running += s
        cum.append(running)
    area = 0.0
    prev = 0.0
    for c in cum:
        area += 0.1 * (prev + c) / 2
        prev = c
    return 1 - 2 * area


TAXED_FUELS = ["coa", "ely", "nga", "oil", "gso", "die", "ker", "lpg"]
CATS = ["app", "che", "clo", "com", "edu", "food", "hea", "hou", "oth", "pap", "pha", "ret", "teq", "tpu"]


def total_direct_effect(decile, sample, stat):
    return sum(direct_effect(f, decile, sample, stat) for f in TAXED_FUELS)


def total_indirect_effect(decile, sample, stat):
    return sum(indirect_effect(c, decile, sample, stat) for c in CATS)


def total_effect(decile, sample, stat):
    return total_direct_effect(decile, sample, stat) + total_indirect_effect(decile, sample, stat)


def grid_row_values(r):
    row = GRID[r]
    return {i: v for i, v in enumerate(row) if v is not None}


if __name__ == "__main__":
    print("=== C.III budget share checks ===")
    print("coa bkt ove mean (expect 0.0):", bshare("coa_share", 0, "Overall", "mean"))
    print("ely bkt ove mean (expect 3.504916...):", bshare("ely_share", 0, "Overall", "mean"))
    print("ely dec ove mean decile1 (expect 3.662860...):", bshare("ely_share", 1, "Overall", "mean"))

    print("\n=== C.V direct effect checks (row 1579 = coa bkt ove mean, expect 0.0) ===")
    print("direct_effect coa bkt ove mean:", direct_effect("coa", 0, "Overall", "mean"))
    # find electricity block header in C.V range for comparison
    for r in range(1576, 1820):
        row = GRID[r]
        if row[2] and "Electricity" in str(row[2]):
            print("C.V electricity header row", r, grid_row_values(r))
            break
    print("direct_effect ely bkt ove mean:", direct_effect("ely", 0, "Overall", "mean"))
    print("direct_effect ely dec ove mean (10 deciles):",
          [direct_effect("ely", d, "Overall", "mean") for d in range(1, 11)])

    print("\n=== C.VII total effect checks (row 2216 = tot bkt ove mean) ===")
    print("grid row 2216:", grid_row_values(2216))
    print("total_effect bkt ove mean:", total_effect(0, "Overall", "mean"))
    print("grid row 2217 (tot dec ove mean):", grid_row_values(2217))
    print("total_effect dec ove mean:", [total_effect(d, "Overall", "mean") for d in range(1, 11)])
