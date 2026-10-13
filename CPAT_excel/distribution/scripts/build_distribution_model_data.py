"""
Exports cpat_excel/Distribution/data_standardized/CPAT_DistributionalData.xlsx
into cpat_data/new_data/distn_*.csv, in the same long/tidy shape as the source
sheets -- this is the bridge between the Excel-side extraction pipeline
(cpat_excel/scripts/*) and the Python model's data convention (InputData /
DATA_PATH, see cpat_model/components/distribution/data.py).

Unlike the forward/reverse pipeline in this folder (which is about *auditing*
that no information is lost converting the raw workbook -> tabular ->
standardized forms), this script's only job is format conversion: one sheet
in, one CSV out, column names unchanged. It intentionally does not pivot,
filter, or rename anything, so the CSVs stay a faithful, inspectable copy of
data_standardized and the modelling logic (which does the pivoting/filtering)
stays in cpat_model where it can be unit tested.

Run from the repo root:
    python cpat_excel/scripts/build_distribution_model_data.py
"""
import os

import openpyxl
import pandas as pd

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SOURCE_XLSX = os.path.join(
    REPO_ROOT, 'cpat_excel', 'Distribution', 'data_standardized', 'CPAT_DistributionalData.xlsx'
)
OUT_DIR = os.path.join(REPO_ROOT, 'cpat_coded', 'cpat_data', 'new_data')

# sheet name -> output CSV file name (without extension)
SHEETS = {
    'IO_GTAP': 'distn_io_gtap',
    'HHSurvey': 'distn_hh_survey',
    'HH_Elast': 'distn_hh_elast',
    'ASPIRE': 'distn_aspire',
    'WHOCooking': 'distn_who_cooking',
    'GDPRatios': 'distn_gdp_ratios',
    'Mapping_SectorCrosswalk': 'distn_mapping_sector_crosswalk',
    'Mapping_CPATSectorsToISIC': 'distn_mapping_cpat_sectors_to_isic',
    'Mapping_IEAFlowsToISIC': 'distn_mapping_iea_flows_to_isic',
    'Mapping_ISICToCPAT': 'distn_mapping_isic_to_cpat',
    'Mapping_CountriesToGTAP10': 'distn_mapping_countries_to_gtap10',
}


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = openpyxl.load_workbook(SOURCE_XLSX, read_only=True, data_only=True)
    for sheet_name, out_name in SHEETS.items():
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        header, data = rows[0], rows[1:]
        df = pd.DataFrame(data, columns=header)
        out_path = os.path.join(OUT_DIR, f'{out_name}.csv')
        df.to_csv(out_path, index=False)
        print(f'{sheet_name} -> {out_path} ({len(df)} rows)')


if __name__ == '__main__':
    main()
