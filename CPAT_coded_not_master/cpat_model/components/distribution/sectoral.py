"""
Step 13 (pseudocode §5) -- sectoral input-cost / output-price outputs.

CPAT Excel: 'Distribution' sheet §D.VI.
"""
import pandas as pd


def rank_affected_sectors(
        price_increase_by_sector: pd.Series,
        hhd_by_sector: pd.Series,
        top_n: int = 20
        ) -> pd.DataFrame:
    """
    Ranks GTAP sectors by their (embedded-energy) output price increase and
    by household-demand share, for the 'top affected sectors' chart pack.

    price_increase_by_sector: Series indexed by GtapSectorCode --
    price_changes.indirect_price_increase_by_gtap_sector output.
    hhd_by_sector: Series indexed by GtapSectorCode, household demand
    (IO_GTAP 'hhd' rows).
    top_n: how many sectors to keep in the 'most affected' cut.

    return: DataFrame indexed by GtapSectorCode, columns price_increase,
    hhd, hhd_share, sorted by price_increase descending; use .head(top_n)
    for the chart-pack cut.
    """
    df = pd.DataFrame({
        'price_increase': price_increase_by_sector,
        'hhd': hhd_by_sector.reindex(price_increase_by_sector.index).fillna(0.0),
    })
    df['hhd_share'] = df['hhd'] / df['hhd'].sum() if df['hhd'].sum() > 0 else 0.0
    return df.sort_values('price_increase', ascending=False).head(top_n)
