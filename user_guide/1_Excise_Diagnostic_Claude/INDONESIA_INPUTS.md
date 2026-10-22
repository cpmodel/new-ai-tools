# Inputs for the Indonesia excise diagnostic

The INPUTS block the master process file asks for (section 2), filled in for Indonesia.
Upload this file to the Claude Project with the master process file.

```yaml
country: "Indonesia"
iso3: "IDN"
currency_code: "IDR; printed in the law as Rp or rupiah"
cut_off_date: "2026-09-30"
run_mode: "WHOLE"            # stages A1-A6 in the master file
period: "2020-01-01 to 2026-09-30"
comparator_report: "none"    # no pre-run comparator yet; omit comparator columns
language: "US English. Short words, short sentences."
```

## Legal texts - a STARTING list for Claude to find and confirm

This list is a pointer, not a source. Claude must find each text, quote it, and give the
link. The reviewer checks every item against the official text. Anything Claude cannot
find is marked data_gap.

- The Excise Law (Law 11 of 1995 on Excise, as amended - including by Law 39 of 2007
  and by the Harmonisation of Tax Regulations Law, Law 7 of 2021).
- Ministry of Finance regulations (PMK) that set the excise rates each year for tobacco
  products, ethyl alcohol and drinks containing ethyl alcohol, for every year in the
  period.
- Vehicle taxes that work like excises: the luxury-goods sales tax (PPnBM) on motor
  vehicles (government regulation and its amendments), and the regional motor vehicle
  taxes (PKB, BBNKB).
- Fuel levies outside the excise law: the regional motor fuel tax (PBBKB) under the
  central-regional fiscal relations law (Law 1 of 2022), and any other fuel charges.
- Announced but not in force at the cut-off (record, do not apply): excise on
  sugar-sweetened beverages (MBDK) and on plastics, as set out in budget documents.

## Budget documents and collections

- Budget law and financial note (APBN, Nota Keuangan) for each year 2020-2026.
- Excise collections by group (tobacco, ethyl alcohol, drinks containing ethyl alcohol)
  for the latest two years: Ministry of Finance (APBN KiTa, LKPP, Directorate General of
  Customs and Excise annual report).
- Nominal GDP in IDR for the same years.

## Cross-check numbers from our dashboard

IMF WoRLD 2026, year 2024 (`TaxMacroFiscalDashboard/build/raw_data/world-imf2026.xlsx`,
the source of the dashboard's tax indicators):

| Indicator | Indonesia |
|---|---|
| Tax revenue (% GDP) | 10.08 |
| Excise revenue (% GDP) | 1.11 |
| Total revenue (% GDP) | 12.84 |

If the diagnostic's collections differ a lot from these, Claude must explain why
(different coverage, year or source).
