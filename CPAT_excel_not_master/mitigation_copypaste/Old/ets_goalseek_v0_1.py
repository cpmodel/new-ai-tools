"""Goal seek for the new-ETS permit price of CPAT-AI-Mitigation-MVP (replaces legacy VBA OverrideETSFast / SolveFast).

The workbook has no VBA and no circular reference. Its section 13 row ets.next proposes the next price for every
year, p x (LN(cap / baseline) / LN(covered / baseline)) ^ exponent (Settings C15), as legacy's
PricesAdjustedByErrorFactor. This script iterates on that proposal with legacy's damped log-space step:

  1. start from the fast estimate (ets.est), or from the override row if it holds values;
  2. step p <- p x (proposal / p) ^ alpha (linear step if a price is 0);
  3. mix with the previous iterate (weight 0.3) from the second iteration;
  4. light smoothing: 5% weight on a centred 3-year moving average;
  5. sanitise: negative or invalid values keep the current price;
  6. alpha halves when the worst error grows (minimum 0.0625) and rises x1.2 when it more than halves (maximum 1);
  7. stop when the worst |covered / cap - 1| over the ETS years is below the tolerance, or after max_iter.

Each iteration writes the prices into the scenario's override row ets.ovr (with MTInputs D_ETSPriceOverride = Yes)
of a working copy, expands the LAMBDAs and recalculates it in LibreOffice. The result is a CSV of prices per year:
paste it into the override row ets.ovr of the scenario and set D_ETSPriceOverride = Yes in MTInputs.

    python ets_goalseek_v0_1.py WORKBOOK.xlsx SCENARIO [SCENARIO ...] [--tol 0.005] [--max-iter 20]
    -> ets_override_s<scenario>.csv next to the workbook
"""
import argparse
import csv
import math
import os

import openpyxl

import build_v1_02 as B
import check_v1_02 as C


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _ma3(v):
    out = []
    for j in range(len(v)):
        w = v[max(j - 1, 0):j + 2]
        out.append(sum(w) / len(w))
    return out


def _step(p, r, alpha):
    if not (math.isfinite(p) and math.isfinite(r)):
        return p
    if p <= 0 or r <= 0:
        return max(p + alpha * (r - p), 0.0)
    return p * (r / p) ** alpha


def _evaluate(wb_path, groups, prices):
    """Write prices into the override rows of a working copy, switch the override on, recalculate.
    Returns {g: (next proposal, gap, cap)} lists over the base year + projection years."""
    wb = openpyxl.load_workbook(wb_path)
    ws, wsm = wb['Mitigation'], wb['MTInputs']
    for g in groups:
        wsm.cell(B.MT_NAMES['D_ETSPriceOverride'], B.MT_COL0 + g - 1).value = 'Yes'
        _, base, years = B.group_cols(g)
        for c, p in zip([base] + years, prices[g]):
            ws.cell(B.R_POL['ets.ovr'], c).value = p
    vals = C.recalc(C.variant(wb, 'expand'))
    out = {}
    for g in groups:
        _, base, years = B.group_cols(g)
        cols = [base] + years
        out[g] = tuple([_num(vals.cell(B.R_EMI[k], c).value) for c in cols] for k in ('ets.next', 'co2.gap',
                                                                                        'co2.cap'))
    return out


def solve(wb_path, groups, tol=0.005, max_iter=12, log=print):
    """Returns ({g: prices base..last year}, {g: worst gap}, iterations)."""
    wb = openpyxl.load_workbook(wb_path)
    first = C.recalc(C.variant(wb, 'expand'))
    p = {}
    for g in groups:
        _, base, years = B.group_cols(g)
        ovr = [wb['Mitigation'].cell(B.R_POL['ets.ovr'], c).value for c in [base] + years]
        p[g] = ([_num(v) for v in ovr] if any(isinstance(v, (int, float)) and v for v in ovr)
                else [_num(first.cell(B.R_POL['ets.est'], c).value) for c in [base] + years])
    alpha, prev_worst, p_prev = {g: 1.0 for g in groups}, {g: float('inf') for g in groups}, {}
    worst = {}
    for it in range(1, max_iter + 1):
        res = _evaluate(wb_path, groups, p)
        done = True
        for g in groups:
            nxt, gap, cap = res[g]
            worst[g] = max([abs(x) for x, k in zip(gap, cap) if k > 0] or [0.0])
            log(f'  iteration {it}, scenario {g}: worst |covered/cap - 1| = {worst[g]:.4f}, alpha {alpha[g]:.3f}, '
                f'price 2030 {p[g][2030 - B.BASE_YEAR]:.2f}, 2035 {p[g][2035 - B.BASE_YEAR]:.2f}')
            if worst[g] < tol:
                continue
            done = False
            new = [_step(a, b, alpha[g]) for a, b in zip(p[g], nxt)]
            if g in p_prev:
                new = [0.7 * a + 0.3 * b for a, b in zip(new, p_prev[g])]
            sm = _ma3(new)
            new = [0.95 * a + 0.05 * b for a, b in zip(new, sm)]
            new = [a if math.isfinite(a) and a >= 0 else b for a, b in zip(new, p[g])]
            if worst[g] > prev_worst[g]:
                alpha[g] = max(alpha[g] * 0.5, 0.0625)
            elif worst[g] < 0.5 * prev_worst[g]:
                alpha[g] = min(alpha[g] * 1.2, 1.0)
            prev_worst[g] = worst[g]
            p_prev[g], p[g] = p[g], new
        if done:
            return p, worst, it
    return p, worst, max_iter


def main():
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('workbook')
    ap.add_argument('scenarios', type=int, nargs='+')
    ap.add_argument('--tol', type=float, default=0.005)
    ap.add_argument('--max-iter', type=int, default=20)
    a = ap.parse_args()
    p, worst, it = solve(a.workbook, a.scenarios, a.tol, a.max_iter)
    years = [B.BASE_YEAR] + B.YEARS
    for g in a.scenarios:
        out = os.path.join(os.path.dirname(os.path.abspath(a.workbook)), f'ets_override_s{g}.csv')
        with open(out, 'w', newline='', encoding='utf-8') as f:
            w = csv.writer(f)
            w.writerow(['year', 'ets_override_price_usd_per_tco2'])
            w.writerows(zip(years, p[g]))
        print(f'scenario {g}: worst gap {worst[g]:.4f} after {it} iterations -> {out}')


if __name__ == '__main__':
    main()
