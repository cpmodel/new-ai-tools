#!/usr/bin/env Rscript
# =============================================================================
# GLOBAL FISCAL INDICATORS — build pipeline (R, all countries)
# =============================================================================
# Mirrors build_ccdr_indicators.py exactly.
#
#   STAGE 1  -> long_all : iso3, indicator, year, value, source   (ALL years)
#   STAGE 2  -> long     : latest observation per economy x indicator
#   STAGE 3  -> wide     : one row per economy, plus matching __year columns
#
# Run:  Rscript build_global_indicators.R          # all 217 countries
#       CCDR_ONLY=1 Rscript build_global_indicators.R   # restrict to the 98
#
# Country universe: every entity in the WDI country metadata with a non-empty
# Region (aggregates have blank Region). is_ccdr flags the 98 CCDR economies.
# Deps: data.table, readxl, writexl, haven
#       install.packages(c("data.table","readxl","writexl","haven"))
#
# -----------------------------------------------------------------------------
# ONE PRE-STEP: the CPAT file is .xlsb, which R cannot read.
#   Either run  python build_global_indicators.py --emit-cpat-xlsx , or
#   open imffossilfuelsubsidiesdata.xlsb in Excel and File > Save As > .xlsx,
#   naming it imffossilfuelsubsidiesdata_sheets.xlsx. Column positions must be
#   preserved - the code below indexes columns by position, not by name.
#
# -----------------------------------------------------------------------------
# WHERE EVERY SOURCE FILE COMES FROM
# -----------------------------------------------------------------------------
# WDI zips - https://data.worldbank.org/indicator/<CODE> -> "Download CSV", or
#   https://api.worldbank.org/v2/country/all/indicator/<CODE>?downloadformat=csv
#   Zip holds 3 CSVs; the data file has 4 header rows to skip and a
#   "Last Updated Date" stamp on line 3.
#     NY.GDP.PETR.RT.ZS oil rents %GDP | NY.GDP.NGAS.RT.ZS gas | NY.GDP.COAL.RT.ZS coal
#     TX.VAL.FUEL.ZS.UN fuel exports %merch | TM.VAL.FUEL.ZS.UN fuel imports %merch
#     EG.IMP.CONS.ZS net energy imports %energy use
# IMF WEO      - https://www.imf.org/external/datamapper/datasets/WEO
#                Wide: COUNTRY x INDICATOR, year cols 2017-2026 (2025-26 projections, dropped).
# IMF WoRLD    - https://www.imf.org/en/topics/fiscal-policies/world-revenue-longitudinal-database
#                Sheet "Data". Revenue variables are ALREADY percent of GDP.
# IMF CPAT/FFS - https://www.imf.org/en/topics/climate-change/energy-subsidies
#                All_Explicit / All_Implicit hold 2024 totals in US$bn (cols 26 / 16).
#                Sheet "data" cols 10..13 are 2021..2024 (1-based in R).
# World Bank IDS - https://databank.worldbank.org/source/international-debt-statistics
#                Country=all, Counterpart-Area=World(WLD), 19 series, Time 2015-2024.
#                ".." = missing. On DT.INR.* a 0 means NO NEW BORROWING, not a 0% rate.
# World Bank BOOST - https://data360.worldbank.org/en/dataset/WB_BOOST
#                PARAM_1 = approved budget, PARAM_2 = executed.
# PEFA         - https://www.pefa.org/assessments/batch-downloads
#                Framework=2016, Type=National, Status=Final, all indicators.
# World Bank PPI - https://ppi.worldbank.org/en/ppidata -> Stata download.
#                Project level, investment US$m, FCY = financial closure year.
# OECD CRDF    - https://www.oecd.org/en/topics/sub-issues/development-finance-for-climate-and-the-environment.html
#                RECIPIENT perspective (not provider).
# Ember        - https://ember-energy.org/data/yearly-electricity-data/
#                No "Oil" category: "Other fossil" = oil plus manufactured gases.
# Steffen 2025 - https://www.nature.com/articles/s41597-025-05912-x
#                Values are strings with "%". NOMINAL after-tax - do not splice with IRENA.
# IRENA / IEA / UNDP - transcribed inline below (PDF tables and country lists).
#                IRENA (2022) Renewable power generation costs in 2021, reproduced as the
#                appendix to IRENA (2023). REFERENCE YEAR 2021, real after-tax.
# =============================================================================

suppressPackageStartupMessages({
  library(data.table); library(readxl); library(writexl); library(haven)
})

# Paths resolve relative to this script unless overridden by environment variable:
#   RAW_DIR=... OUT_DIR=... CONFIG_DIR=... Rscript build_global_indicators.R
.args <- commandArgs(trailingOnly = FALSE)
.here <- dirname(sub("^--file=", "", .args[grep("^--file=", .args)][1]))
if (is.na(.here) || !nzchar(.here)) .here <- getwd()
RAW    <- normalizePath(Sys.getenv("RAW_DIR",    file.path(.here, "..", "raw_data")), mustWork = FALSE)
OUT    <- normalizePath(Sys.getenv("OUT_DIR",    file.path(.here, "..", "output")),   mustWork = FALSE)
CONFIG <- normalizePath(Sys.getenv("CONFIG_DIR", file.path(.here, "..", "config")),   mustWork = FALSE)
dir.create(OUT, showWarnings = FALSE, recursive = TRUE)
CPAT_XLSX <- file.path(OUT, "imffossilfuelsubsidiesdata_sheets.xlsx")  # from --emit-cpat-xlsx

# -----------------------------------------------------------------------------
# CCDR economy list (98 published Country Climate and Development Reports)
# -----------------------------------------------------------------------------
CCDR_ONLY <- nzchar(Sys.getenv("CCDR_ONLY"))

# --- country universe from the WDI country metadata --------------------------
.mp   <- list.files(RAW, pattern = glob2rx("API_*_DS2_*.zip"), full.names = TRUE)[1]
.inner<- grep("^Metadata_Country", unzip(.mp, list = TRUE)$Name, value = TRUE)[1]
meta  <- as.data.table(read.csv(unz(.mp, .inner), check.names = FALSE, encoding = "UTF-8"))
# The WDI metadata CSV carries a UTF-8 BOM, so the first column name arrives as
# "<BOM>Country Code". Strip non-alphanumerics from names before referencing them.
nm <- gsub("[^A-Za-z]", "", names(meta))
nm[nm == ""] <- paste0("V", which(nm == ""))
setnames(meta, nm)
real  <- meta[!is.na(Region) & trimws(Region) != ""]
NAME   <- setNames(real$TableName,   real$CountryCode)
REGION <- setNames(real$Region,      real$CountryCode)
INCOME <- setNames(real$IncomeGroup, real$CountryCode)
ISO    <- real$CountryCode

CCDR_ISO <- strsplit(gsub("[\n ]+", " ", paste(
"AGO BWA COM COD ETH KEN MDG MWI MUS MOZ NAM RWA SYC SOM ZAF SSD TZA UGA ZMB ZWE BEN BFA CPV CMR",
"CAF TCD COG CIV GAB GMB GHA GIN GNB LBR MLI MRT NER SEN SLE TGO KHM CHN FJI IDN KIR LAO MYS MHL",
"MNG PHL THA TUV VNM ALB ARM AZE BIH HRV GEO KAZ XKX KGZ MDA MNE MKD POL ROU SRB TJK TUR UZB ARG",
"BRA COL DMA DOM ECU GRD HND PRY PER LCA VCT DJI EGY IRQ JOR LBN MAR PAK TUN PSE YEM BGD BTN MDV",
"NPL LKA")), " ")[[1]]
if (CCDR_ONLY) ISO <- intersect(ISO, CCDR_ISO)

# --- name -> ISO3 matching ----------------------------------------------------
# Sources spell countries differently. Strip accents, parentheticals and
# administrative suffixes, then fall back to the text before the first comma.
# Aggregates deliberately fail to match and are dropped.
SUFF <- c(paste0(",?\\s*(the\\s+)?(islamic|arab|bolivarian|plurinational|federal democratic|",
                 "federated|people s|democratic|socialist|principality|grand duchy)?\\s*",
                 "(republic|kingdom|state|states|union|commonwealth|federation|principality)s?",
                 "\\s+of(\\s+the)?$"),
          ",\\s*(rep\\.|dem\\. rep\\.|the|fed\\. sts\\.|islamic rep\\.|arab rep\\.|rb|prc|cr)$",
          "^the\\s+")
# iconv(to = "ASCII//TRANSLIT") depends on locale: under C/POSIX it emits "?"
# instead of transliterating, which silently breaks every accented name. Use an
# explicit character table instead - deterministic everywhere.
.FROM <- "\u00e0\u00e1\u00e2\u00e3\u00e4\u00e5\u00e8\u00e9\u00ea\u00eb\u00ec\u00ed\u00ee\u00ef\u00f2\u00f3\u00f4\u00f5\u00f6\u00f9\u00fa\u00fb\u00fc\u00fd\u00f1\u00e7\u00c0\u00c1\u00c2\u00c3\u00c4\u00c5\u00c8\u00c9\u00ca\u00cb\u00cc\u00cd\u00ce\u00cf\u00d2\u00d3\u00d4\u00d5\u00d6\u00d9\u00da\u00db\u00dc\u00dd\u00d1\u00c7"
.TO   <- "aaaaaaeeeeiiiiooooouuuuyncAAAAAAEEEEIIIIOOOOOUUUUYNC"
.deacc <- function(x) {
  x <- chartr(.FROM, .TO, enc2utf8(as.character(x)))
  x <- gsub("\u00f8", "o", x); x <- gsub("\u00d8", "O", x)
  x <- gsub("\u00e6", "ae", x); x <- gsub("\u00df", "ss", x)
  x
}
.nrm <- function(s) {
  s <- tolower(trimws(ifelse(is.na(s), "", .deacc(s))))
  s <- gsub("\\([^)]*\\)", "", s)
  for (i in 1:3) for (p in SUFF) s <- trimws(gsub(p, "", s))
  trimws(gsub("[^a-z0-9]+", " ", s))
}
ALIAS <- c("vietnam"="VNM","viet nam"="VNM","turkiye"="TUR","turkey"="TUR","egypt"="EGY",
 "gambia"="GMB","yemen"="YEM","kyrgyzstan"="KGZ","lao pdr"="LAO","laos"="LAO",
 "lao people s democratic republic"="LAO","cape verde"="CPV","cabo verde"="CPV",
 "west bank and gaza strip"="PSE","cote d ivoire"="CIV","ivory coast"="CIV",
 "st lucia"="LCA","saint lucia"="LCA","st vincent and grenadines"="VCT",
 "saint vincent and grenadines"="VCT","saint vincent and the grenadines"="VCT",
 "st vincent and the grenadines"="VCT","brunei"="BRN","russia"="RUS","syria"="SYR",
 "iran"="IRN","venezuela"="VEN","bolivia"="BOL","tanzania"="TZA","moldova"="MDA",
 "north macedonia"="MKD","macedonia"="MKD","micronesia"="FSM","bahamas"="BHS",
 "hong kong"="HKG","hong kong special administrative region"="HKG","macao"="MAC",
 "macao special administrative region"="MAC","korea"="KOR","south korea"="KOR",
 "republic of korea"="KOR","north korea"="PRK","democratic people s republic of korea"="PRK",
 "slovak"="SVK","czechia"="CZE","czech"="CZE","czech republic"="CZE","eswatini"="SWZ",
 "swaziland"="SWZ","timor leste"="TLS","east timor"="TLS","kosovo"="XKX","curacao"="CUW",
 "chinese taipei"="TWN","taiwan"="TWN","taiwan province of china"="TWN","china"="CHN",
 "somalia"="SOM","guyana"="GUY","st kitts and nevis"="KNA","saint kitts and nevis"="KNA",
 "cook islands"="COK","nauru"="NRU","niue"="NIU","bosnia and herzegovina"="BIH",
 "montserrat"="MSR","saint helena"="SHN","tokelau"="TKL","wallis and futuna"="WLF",
 "st pierre and miquelon"="SPM","saint pierre and miquelon"="SPM")
.LUT <- c(setNames(names(NAME), .nrm(unname(NAME))), ALIAS)
UNMATCHED <- list()
iso_of <- function(x) {
  raw <- as.character(x)
  # Guard: stripping "Democratic Republic of the" reduces DR Congo to "congo",
  # colliding with Republic of Congo. Decide these two before normalising.
  rl  <- tolower(ifelse(is.na(raw), "", raw))
  out <- rep(NA_character_, length(raw))
  cg  <- grepl("congo", rl)
  out[cg] <- ifelse(grepl("dem|drc|kinshasa", rl[cg]), "COD", "COG")
  for (f in list(function(z) .nrm(z),
                 function(z) .nrm(sub(",.*$", "", z)),
                 function(z) .nrm(sub(" - .*$", "", z)),
                 function(z) trimws(gsub("[^a-z0-9]+", " ", tolower(.deacc(z)))))) {
    i <- is.na(out)
    if (!any(i)) break
    out[i] <- unname(.LUT[f(raw[i])])
  }
  out
}
to_iso <- function(x, source) {
  out <- iso_of(x)
  miss <- sort(unique(as.character(x)[is.na(out) & !is.na(x)]))
  if (length(miss)) UNMATCHED[[source]] <<- miss
  out
}

ACC <- list()   # accumulator of long chunks
emit <- function(iso3, indicator, year, value, source) {
  d <- data.table(iso3 = as.character(iso3), indicator = indicator,
                  year = as.character(year), value = suppressWarnings(as.numeric(value)),
                  source = source)
  d <- d[!is.na(iso3) & iso3 %in% ISO & !is.na(value)]
  if (nrow(d)) ACC[[length(ACC) + 1L]] <<- d
}

# read a WDI download zip -> long (iso3, year, value); 4 header rows
wdi_zip <- function(pattern) {
  path <- list.files(RAW, pattern = glob2rx(pattern), full.names = TRUE)[1]
  inner <- grep("^API_.*csv$", unzip(path, list = TRUE)$Name, value = TRUE)
  inner <- inner[!grepl("Metadata", inner)][1]
  con <- unz(path, inner)
  df  <- as.data.table(read.csv(con, skip = 4, check.names = FALSE, encoding = "UTF-8"))
  yrs <- grep("^[0-9]{4}$", names(df), value = TRUE)
  m <- melt(df[, c("Country Code", yrs), with = FALSE], id.vars = "Country Code",
            variable.name = "year", value.name = "value", variable.factor = FALSE)
  setnames(m, "Country Code", "iso3")
  list(data = m[!is.na(value)], file = basename(path))
}

# =============================================================================
# STAGE 1 — load every source into one long frame (ALL years retained)
# =============================================================================

## --- IMF WEO ----------------------------------------------------------------
WEO_F <- list.files(RAW, pattern = "^dataset_.*IMF_RES_WEO.*csv$", full.names = TRUE)[1]
weo <- as.data.table(read.csv(WEO_F, check.names = FALSE, encoding = "UTF-8"))
weo[, iso := to_iso(COUNTRY, "IMF WEO")]
WY <- grep("^[0-9]{4}$", names(weo), value = TRUE); WY <- WY[as.integer(WY) <= 2024]
wf <- basename(WEO_F)
# Two WEO labels can resolve to one ISO; keep the first non-null per ISO.
weo_ser <- function(ind) weo[INDICATOR == ind & !is.na(iso), c("iso", WY), with = FALSE][
                            , lapply(.SD, function(z) z[!is.na(z)][1]), by = iso]

G  <- "Gross debt, General government, Percent of GDP"
R_ <- "Revenue, General government, Percent of GDP"
NB <- "Net lending (+) / net borrowing (-), General government, Percent of GDP"
PB <- "Primary net lending (+) / net borrowing (-), General government, Percent of GDP"
GDPU <- "Gross domestic product (GDP), Current prices, US dollar"

gd <- melt(weo_ser(G), id.vars = "iso", variable.name = "year", value.name = "v",
           variable.factor = FALSE)
emit(gd$iso, "gross_debt_pct_gdp", gd$year, gd$v, wf)

# interest = primary balance less overall balance; then divided by revenue
mk <- function(ind, nm) melt(weo_ser(ind), id.vars = "iso", variable.name = "year",
                             value.name = nm, variable.factor = FALSE)
ir <- Reduce(function(a, b) merge(a, b, by = c("iso", "year")),
             list(mk(PB, "pb"), mk(NB, "nb"), mk(R_, "rev")))
ir <- ir[!is.na(pb) & !is.na(nb) & !is.na(rev) & rev > 0][, itr := (pb - nb) / rev * 100][
        itr >= 0 & itr < 100]
emit(ir$iso, "interest_to_revenue_pct", ir$year, ir$itr, paste(wf, "(derived)"))

gg <- melt(weo_ser(GDPU), id.vars = "iso", variable.name = "year", value.name = "gdp",
           variable.factor = FALSE)[!is.na(gdp)]
setorder(gg, iso, year); GDPdt <- gg[, .SD[.N], by = iso]      # latest actual year
GDP <- setNames(GDPdt$gdp, GDPdt$iso)

## --- IMF WoRLD (values already % of GDP) -------------------------------------
WORLD_F <- "world-imf2026.xlsx"
wo <- as.data.table(read_excel(file.path(RAW, WORLD_F), sheet = "Data"))
wo <- wo[ISO3 %in% ISO & year >= 2010 & year <= 2024]
emit(wo$ISO3, "tax_revenue_pct_gdp", wo$year, wo$TaxRev, WORLD_F)
we <- wo[!is.na(TaxSalExc) & !is.na(TotRev) & TotRev > 0]
emit(we$ISO3, "excise_revenue_pct_of_revenue", we$year, we$TaxSalExc / we$TotRev * 100, WORLD_F)

## --- IMF CPAT / Fossil Fuel Subsidies (via the converted .xlsx) --------------
FFS_F <- "imffossilfuelsubsidiesdata.xlsb"
stopifnot("run convert_cpat_xlsb.py first (see pre-step note above)" = file.exists(CPAT_XLSX))
rd <- function(sh) as.data.table(read_excel(CPAT_XLSX, sheet = sh))  # headers c0..cN
expl <- rd("All_Explicit"); impl <- rd("All_Implicit"); cdat <- rd("data")
ffs_tot <- function(dt, total_col) {                 # c1 = iso3, total_col = the 'Total' block
  x <- data.table(iso = toupper(trimws(as.character(dt[["c1"]]))),
                  v   = suppressWarnings(as.numeric(dt[[total_col]])))
  x[nchar(iso) == 3 & !is.na(v)]
}
e <- ffs_tot(expl, "c25"); i_ <- ffs_tot(impl, "c15")
gdp_dt <- data.table(iso = names(GDP), gdp = as.numeric(GDP))
e2 <- merge(e, gdp_dt, by = "iso"); i2 <- merge(i_, gdp_dt, by = "iso")
emit(e2$iso, "ff_subsidy_explicit_pct_gdp", 2024, e2$v / e2$gdp * 100, FFS_F)
emit(i2$iso, "ff_subsidy_implicit_pct_gdp", 2024, i2$v / i2$gdp * 100, FFS_F)
t2 <- merge(e2[, .(iso, e = v, gdp)], i2[, .(iso, i = v)], by = "iso")
emit(t2$iso, "ff_subsidy_total_pct_gdp", 2024, (t2$e + t2$i) / t2$gdp * 100,
     paste(FFS_F, "+ WEO"))
# renewables share: indicator name in col 3, scenario in col 2, iso in col 5, 2021-24 in 10:13
ren <- cdat[c2 == "Share of electricity from renewables - baseline" & c1 == "U1"]
if (nrow(ren)) for (k in 1:4) {
  v <- suppressWarnings(as.numeric(ren[[paste0("c", 8 + k)]]))   # c9..c12 = 2021..2024
  v <- ifelse(!is.na(v) & v <= 1.0001, v * 100, v)
  emit(toupper(as.character(ren$c4)), "renewable_share_electricity_pct", 2020 + k, v, FFS_F)
}
# Retail (consumer) prices. Diesel/gasoline $/litre; gas and coal $/GJ.
# CPAT has no COMMERCIAL sector - it splits residential / industry / power / other.
CPAT_PRICE <- c("Retail price, die, all"="price_diesel_usd_litre",
 "Retail price, gso, all"="price_gasoline_usd_litre",
 "Retail price, nga, res"="price_gas_residential_usd_gj",
 "Retail price, nga, ind"="price_gas_industry_usd_gj",
 "Retail price, nga, pow"="price_gas_power_usd_gj",
 "Retail price, coa, res"="price_coal_residential_usd_gj",
 "Retail price, coa, ind"="price_coal_industry_usd_gj",
 "Retail price, coa, pow"="price_coal_power_usd_gj",
 "Retail price, ecy, res"="price_electricity_residential_usd_kwh",
 "Retail price, ecy, ind"="price_electricity_industry_usd_kwh")
for (nm in names(CPAT_PRICE)) {
  pr <- cdat[c2 == nm & c1 == "U1"]
  if (!nrow(pr)) next
  for (k in 1:4) {
    v <- suppressWarnings(as.numeric(pr[[paste0("c", 8 + k)]]))
    v[!is.na(v) & v <= 0] <- NA
    emit(toupper(as.character(pr$c4)), unname(CPAT_PRICE[nm]), 2020 + k, v, FFS_F)
  }
}

## --- World Bank IDS ----------------------------------------------------------
IDS_F <- "P_Data_Extract_From_International_Debt_Statistics__2_.xlsx"
ids <- as.data.table(read_excel(file.path(RAW, IDS_F), sheet = "Data"))
setnames(ids, make.names(names(ids), unique = TRUE))
ycols <- grep("^X20[0-9]{2}", names(ids), value = TRUE)
ycols <- ycols[as.integer(substr(ycols, 2, 5)) <= 2024]
ids <- ids[!is.na(Country.Code) & Country.Code %in% ISO]
for (y in ycols) set(ids, j = y, value = suppressWarnings(as.numeric(ids[[y]])))
ids_panel <- function(code, nonzero = FALSE) {
  d <- ids[Series.Code == code, c("Country.Code", ycols), with = FALSE]
  m <- melt(d, id.vars = "Country.Code", variable.name = "yc", value.name = "v",
            variable.factor = FALSE)
  m[, year := substr(yc, 2, 5)][!is.na(v)][if (nonzero) v > 0 else TRUE][
    , .(iso3 = Country.Code, year, v)]
}
u <- ids_panel("DT.CUR.USDL.ZS");  emit(u$iso3, "ids_usd_share_ppg_debt_pct", u$year, u$v, IDS_F)
mt <- ids_panel("DT.MAT.DPPG", TRUE); emit(mt$iso3, "ids_avg_maturity_years", mt$year, mt$v, IDS_F)
pr <- ids_panel("DT.INR.PRVT", TRUE); emit(pr$iso3, "ids_private_creditor_rate_pct", pr$year, pr$v, IDS_F)
of <- ids_panel("DT.INR.OFFT", TRUE)
setorder(pr, iso3, year); setorder(of, iso3, year)
sp <- merge(pr[, .SD[.N], by = iso3][, .(iso3, yp = year, p = v)],
            of[, .SD[.N], by = iso3][, .(iso3, yo = year, o = v)], by = "iso3")
emit(sp$iso3, "ids_official_private_rate_spread_pp", pmax(sp$yp, sp$yo), sp$p - sp$o, IDS_F)
tot <- ids_panel("DT.DOD.DPPG.CD")
for (cd in list(c("DT.DOD.PRVT.CD", "ids_private_creditor_share_pct"),
                c("DT.DOD.MLAT.CD", "ids_multilateral_share_pct"))) {
  pt <- merge(ids_panel(cd[1]), tot, by = c("iso3", "year"), suffixes = c("", ".t"))
  pt <- pt[v.t > 0]
  emit(pt$iso3, cd[2], pt$year, pt$v / pt$v.t * 100, IDS_F)
}

## --- World Bank BOOST --------------------------------------------------------
BOOST_F <- "WB_BOOST_WIDEF.csv"
bo <- fread(file.path(RAW, BOOST_F), encoding = "UTF-8")
BY <- grep("^[0-9]{4}$", names(bo), value = TRUE)
cap <- bo[INDICATOR == "WB_BOOST_EXP_ECON_CAP_EXP" & UNIT_MEASURE == "XDC"]
mlt <- function(d) melt(d[, c("REF_AREA", BY), with = FALSE], id.vars = "REF_AREA",
                        variable.name = "year", value.name = "v", variable.factor = FALSE)
A <- mlt(cap[COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_1"])[!is.na(v) & v > 0]
E <- mlt(cap[COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_2"])[!is.na(v)]
ex <- merge(A, E, by = c("REF_AREA", "year"), suffixes = c(".a", ".e"))
emit(ex$REF_AREA, "capex_execution_rate_pct", ex$year, ex$v.e / ex$v.a * 100, BOOST_F)
en <- mlt(bo[INDICATOR == "WB_BOOST_EXP_FUNC_ENV_PRO" & UNIT_MEASURE == "PT_EXP" &
             COMP_BREAKDOWN_1 == "WB_BOOST_BUDGET_PARAM_2"])[!is.na(v)]
emit(en$REF_AREA, "env_protection_pct_of_expenditure", en$year, en$v, BOOST_F)

## --- PEFA --------------------------------------------------------------------
PEFA_F <- "assessments_1787749571.csv"
pe <- fread(file.path(RAW, PEFA_F), encoding = "UTF-8")
SCORE <- c("A"=4, "B+"=3.5, "B"=3, "C+"=2.5, "C"=2, "D+"=1.5, "D"=1, "D*"=1)
pe[, iso := to_iso(Country, "PEFA")]
for (cd in list(c("PI-11","pefa_pi11_pim_score"), c("PI-12.2","pefa_pi122_asset_monitoring"),
                c("PI-10.3","pefa_pi103_contingent_liabilities"))) {
  d <- pe[!is.na(iso) & get(cd[1]) %in% names(SCORE)]
  emit(d$iso, cd[2], d$Year, unname(SCORE[d[[cd[1]]]]), PEFA_F)
}

## --- World Bank PPI (5-year average -> one derived observation) ---------------
PPI_F <- "2024-PPI-Full-DTA.dta"
pp <- as.data.table(read_dta(file.path(RAW, PPI_F)))
pp[, cname := as.character(haven::as_factor(country))][, iso := to_iso(cname, "World Bank PPI")]
ag <- pp[FCY >= 2020 & FCY <= 2024 & !is.na(iso), .(v = sum(investment, na.rm = TRUE)/5/1000), by = iso]
ag <- merge(ag, gdp_dt, by = "iso")[v > 0]
emit(ag$iso, "ppi_commitments_pct_gdp", "2020-2024 avg", ag$v / ag$gdp * 100,
     paste(PPI_F, "+ WEO"))

## --- OECD CRDF, recipient perspective ----------------------------------------
CRDF_F <- "CRDF-RP_2024.xlsx"
cr <- as.data.table(read_excel(file.path(RAW, CRDF_F), sheet = "2024"))
MITCOL <- "Mitigation-related development finance (includes overlap) - Commitment - 2024 USD thousand"
cr[, iso := to_iso(`Recipient Name`, "OECD CRDF")]
mf <- cr[!is.na(iso), .(v = sum(get(MITCOL), na.rm = TRUE)/1e6), by = iso]
mf <- merge(mf, gdp_dt, by = "iso")[v > 0]
emit(mf$iso, "mitigation_finance_received_pct_gdp", 2024, mf$v / mf$gdp * 100,
     paste(CRDF_F, "+ WEO"))

## --- Ember --------------------------------------------------------------------
EMBER_F <- "release_generation_yearly_global.csv"
em <- fread(file.path(RAW, EMBER_F), encoding = "UTF-8")
em_all <- em[`Area type` == "Country or economy"]
of <- em_all[`Electricity source` == "Other fossil" & !is.na(`Share of generation (%)`)]
emit(of$`ISO 3 code`, "other_fossil_gen_share_pct", of$Year, of$`Share of generation (%)`, EMBER_F)
# Ember publishes the YoY share change directly; capacity growth is derived from GW.
ren_e <- em_all[`Electricity source` == "Renewables"]
# Ember carries a country's record forward unchanged where it has no observations
# (e.g. Kiribati: identical 0.04 TWh, 25% share, every year 2020-24). The YoY change
# then reads a spurious 0.0 pp. Detect it by the WHOLE record being constant - total
# generation flat as well as the share - which distinguishes carry-forward from a
# genuinely stable mix like Iceland or Paraguay.
recent <- em_all[Year >= max(Year, na.rm = TRUE) - 5]
tg <- recent[`Electricity source` == "Total generation",
             .(n = uniqueN(`Generation (TWh)`)), by = `ISO 3 code`]
rs <- recent[`Electricity source` == "Renewables",
             .(n = uniqueN(`Share of generation (%)`)), by = `ISO 3 code`]
carry <- merge(rs, tg, by = "ISO 3 code", all.x = TRUE, suffixes = c(".r", ".t"))
CARRY_FWD <- carry[n.r <= 1 & (is.na(n.t) | n.t <= 1)]$`ISO 3 code`
sh_ <- ren_e[!is.na(`Share of generation YoY change (% points)`) &
             !(`ISO 3 code` %in% CARRY_FWD)]
emit(sh_$`ISO 3 code`, "re_share_change_pp", sh_$Year,
     sh_$`Share of generation YoY change (% points)`, EMBER_F)
cp <- ren_e[!is.na(`Capacity (GW)`), .(iso3 = `ISO 3 code`, Year, cap = `Capacity (GW)`)]
setorder(cp, iso3, Year)
cp[, prev := shift(cap), by = iso3]
cp <- cp[!is.na(prev) & prev > 0]
emit(cp$iso3, "re_capacity_growth_pct", cp$Year, (cp$cap / cp$prev - 1) * 100, EMBER_F)

## --- WDI energy and rents -----------------------------------------------------
wdi_map <- list(c("API_NY_GDP_PETR_RT_ZS*.zip","oil_rents_pct_gdp"),
                c("API_NY_GDP_NGAS_RT_ZS*.zip","gas_rents_pct_gdp"),
                c("API_NY_GDP_COAL_RT_ZS*.zip","coal_rents_pct_gdp"),
                c("API_TX_VAL_FUEL_ZS_UN*.zip","fuel_exports_pct_merch"),
                c("API_TM_VAL_FUEL_ZS_UN*.zip","fuel_imports_pct_merch"),
                c("API_EG_IMP_CONS_ZS*.zip","energy_import_dep_pct"))
for (m in wdi_map) { z <- wdi_zip(m[1]); emit(z$data$iso3, m[2], z$data$year, z$data$value, z$file) }

## --- World Bank RISE ----------------------------------------------------------
RISE_F <- "WB_RISE_WIDEF.csv"
ri <- fread(file.path(RAW, RISE_F), encoding = "UTF-8")
RY <- grep("^[0-9]{4}$", names(ri), value = TRUE)
RISE_KEEP <- c("WB_RISE_RE_ALL"="rise_renewable_energy_score",
               "WB_RISE_RE_LVL_PLYNG_FLD"="rise_level_playing_field_score",
               "WB_RISE_EE_FMEE"="rise_ee_financing_mechanisms_score")
for (code in names(RISE_KEEP)) {
  d <- ri[INDICATOR == code, c("REF_AREA", RY), with = FALSE]
  m <- melt(d, id.vars = "REF_AREA", variable.name = "year", value.name = "v",
            variable.factor = FALSE)[!is.na(v)]
  emit(m$REF_AREA, unname(RISE_KEEP[code]), m$year, m$v, RISE_F)
}

## --- Power Sector Reform Tracker (PSRT) ---------------------------------------
## Market structure coded from the reform flags:
##   3 = wholesale market   2 = single buyer (IPPs, no market)   1 = vertically integrated
PSRT_F <- "dataverse_files.zip"
.td <- tempfile(); dir.create(.td)
.dta <- grep("PSRT_V1_Database\\.dta$", unzip(file.path(RAW, PSRT_F), list = TRUE)$Name, value = TRUE)[1]
unzip(file.path(RAW, PSRT_F), files = .dta, exdir = .td)
ps <- as.data.table(read_dta(file.path(.td, .dta)))
ps[, iso := to_iso(cntry, "PSRT")]
ps <- ps[!is.na(iso)]
emit(ps$iso, "power_market_structure", ps$year,
     fifelse(ps$r_wem == 1, 3, fifelse(ps$r_ipp == 1, 2, 1)), PSRT_F)
emit(ps$iso, "independent_regulator", ps$year, as.integer(ps$r_reg), PSRT_F)
emit(ps$iso, "generation_unbundled",  ps$year, as.integer(ps$r_und), PSRT_F)

## --- WACC: IRENA (transcribed), Steffen, IEA ----------------------------------
IRENA_SOLAR <- c(DZA=11.0,ARG=13.8,AZE=7.0,BGD=6.8,BIH=10.4,BRA=6.3,BFA=5.8,CHN=2.5,COL=5.6,
 HRV=5.3,DOM=5.6,ECU=12.2,EGY=8.8,ETH=8.4,GHA=9.5,HND=4.6,IDN=6.0,IRQ=9.6,JOR=8.2,KAZ=6.3,
 KEN=8.4,LBN=21.0,MYS=5.4,MUS=4.6,MNG=8.0,MNE=8.8,MAR=6.7,NAM=4.2,PAK=9.2,PER=5.2,PHL=5.7,
 POL=3.9,ROU=5.1,RWA=5.6,SEN=4.3,ZAF=5.2,LKA=10.3,THA=4.5,TUN=9.3,TUR=7.5,UGA=6.9,VNM=6.0,
 YEM=17.2)
emit(names(IRENA_SOLAR), "wacc_solar_real_aftertax_pct", 2021, unname(IRENA_SOLAR),
     "IRENA_Cost_of_financing_renewable_power_Appendix_2023.pdf")
STEF_F <- "SteffenEtAl2025_WACC_database.csv"
st <- fread(file.path(RAW, STEF_F), encoding = "UTF-8")
st <- st[Variable == "WACC (nominal after-tax)"]
st[, val := suppressWarnings(as.numeric(gsub("%", "", trimws(as.character(Value)))))]
st <- st[!is.na(val) & ISO3 %in% ISO]
sm <- st[, .(ly = max(`Financing year`)), by = ISO3]
sm <- merge(st, sm, by.x = c("ISO3","Financing year"), by.y = c("ISO3","ly"))[
       , .(v = mean(val)), by = .(ISO3, `Financing year`)]
emit(sm$ISO3, "wacc_steffen_nominal_aftertax_pct", sm$`Financing year`, sm$v, STEF_F)
IEA_COUNTRIES <- c("BRA","IDN","KEN","MYS","PHL","SEN","THA","VNM","ZAF")   # CofCObservatoryData.xlsx
emit(IEA_COUNTRIES, "wacc_iea_observatory_available", 2024, 1, "CofCObservatoryData.xlsx")

## --- UNDP climate-PFM diagnostic flag (Table 4.2, transcribed) ----------------
UNDP_YES <- c("ARM","BGD","BEN","KHM","CHN","COL","ECU","ETH","FJI","GHA","HND","IDN","KEN",
 "KIR","MHL","MAR","MOZ","NPL","PAK","PHL","RWA","SYC","TZA","THA","UGA","VNM","AZE","GEO",
 "HRV","GMB","MDG","MUS","PER")
emit(ISO, "climate_pfm_diagnostic_on_record", 2022, as.integer(ISO %in% UNDP_YES),
     "UNDP-Global-Climate-Public-Finance-Review-2022.pdf (Table 4.2)")

long_all <- rbindlist(ACC, use.names = TRUE)
long_all[, value := round(value, 3)]
long_all[, `:=`(country = unname(NAME[iso3]), region = unname(REGION[iso3]),
                is_ccdr = iso3 %in% CCDR_ISO)]
setcolorder(long_all, c("iso3","country","region","is_ccdr","indicator","year","value","source"))

# =============================================================================
# STAGE 2 — THE FILTER: latest observation per economy x indicator
# =============================================================================
long <- long_all[order(iso3, indicator, suppressWarnings(as.integer(substr(year, 1, 4))))][
          , .SD[.N], by = .(iso3, indicator)]

## --- derived indicators that combine two filtered series ----------------------
pick <- function(ind) long[indicator == ind, .(iso3, v = value, y = year)]
o <- pick("oil_rents_pct_gdp"); g_ <- pick("gas_rents_pct_gdp"); c_ <- pick("coal_rents_pct_gdp")
fr <- merge(merge(o, g_, by = "iso3", suffixes = c(".o", ".g")), c_, by = "iso3")
add <- list(data.table(iso3 = fr$iso3, indicator = "fossil_rents_total_pct_gdp",
                       year = fr$y.o, value = fr$v.o + fr$v.g + fr$v,
                       source = "WDI NY.GDP.PETR+NGAS+COAL.RT.ZS"))
fx <- pick("fuel_exports_pct_merch"); fm <- pick("fuel_imports_pct_merch")
nf <- merge(fx, fm, by = "iso3", suffixes = c(".x", ".m"))
add[[2]] <- data.table(iso3 = nf$iso3, indicator = "net_fuel_trade_pct_merch",
                       year = nf$y.x, value = nf$v.x - nf$v.m,
                       source = "WDI TX.VAL.FUEL.ZS.UN less TM.VAL.FUEL.ZS.UN")
wsrc <- c("wacc_solar_real_aftertax_pct","wacc_steffen_nominal_aftertax_pct",
          "wacc_iea_observatory_available")
wn <- long[indicator %in% wsrc, .N, by = iso3]
wn <- merge(data.table(iso3 = ISO), wn, by = "iso3", all.x = TRUE)[is.na(N), N := 0L]
# 0 sources means no WACC evidence at all - a missing row, not a value of 0
add[[3]] <- data.table(iso3 = wn[N > 0]$iso3, indicator = "wacc_n_sources", year = "2021-2024",
                       value = as.numeric(wn[N > 0]$N), source = "IRENA / Steffen et al. / IEA")
long <- rbindlist(list(long[, .(iso3, indicator, year, value, source)], rbindlist(add)),
                  use.names = TRUE)

## --- keep only the 35 dashboard indicators, in block order --------------------
# The indicator set is defined ONCE, in config/indicators.csv, shared with the
# Python pipeline and with build_dashboard.py.
.spec <- fread(file.path(CONFIG, "indicators.csv"), encoding = "UTF-8")
setorder(.spec, order)
KEEP <- .spec$code

long <- long[indicator %in% KEEP]
long[, `:=`(country = unname(NAME[iso3]), region = unname(REGION[iso3]),
            income = unname(INCOME[iso3]), is_ccdr = iso3 %in% CCDR_ISO)]
long[, indicator := factor(indicator, levels = KEEP)]
setorder(long, country, indicator)
long <- long[, .(iso3, country, region, income, is_ccdr, indicator, value, year, source)]

# =============================================================================
# STAGE 3 — wide (dashboard layout: values then matching __year columns)
# =============================================================================
wv <- dcast(long, iso3 ~ indicator, value.var = "value")
wy <- dcast(long, iso3 ~ indicator, value.var = "year")
setnames(wy, setdiff(names(wy), "iso3"), paste0(setdiff(names(wy), "iso3"), "__year"))
wide <- merge(wv, wy, by = "iso3")
wide[, `:=`(country = unname(NAME[iso3]), region = unname(REGION[iso3]),
            income = unname(INCOME[iso3]), is_ccdr = iso3 %in% CCDR_ISO)]
setcolorder(wide, c("iso3","country","region","income","is_ccdr", KEEP, paste0(KEEP, "__year")))
setorder(wide, country)

# =============================================================================
# OUTPUT
# =============================================================================
PFX <- if (CCDR_ONLY) "ccdr_" else "global_"
write_xlsx(list(Long_all_years = as.data.frame(long_all),
                Data_long      = as.data.frame(long),
                Data_wide      = as.data.frame(wide)),
           file.path(OUT, paste0(PFX, "indicators_R.xlsx")))
fwrite(long_all, file.path(OUT, paste0(PFX, "long_all_years_R.csv")))
fwrite(long,     file.path(OUT, paste0(PFX, "data_long_R.csv")))
fwrite(wide,     file.path(OUT, paste0(PFX, "data_wide_R.csv")))

cat(sprintf("Stage 1  long_all : %6s rows  %d series  %d economies\n",
            format(nrow(long_all), big.mark = ","),
            uniqueN(long_all$indicator), uniqueN(long_all$iso3)))
cat(sprintf("Stage 2  long     : %6s rows  %d indicators\n",
            format(nrow(long), big.mark = ","), uniqueN(long$indicator)))
cat(sprintf("Stage 3  wide     : %d economies x %d indicators\n", nrow(wide), length(KEEP)))
miss <- setdiff(KEEP, unique(as.character(long$indicator)))
if (length(miss)) stop(sprintf(
  "PIPELINE/CONFIG MISMATCH - config/indicators.csv lists indicators this pipeline never emits:\n   %s",
  paste(miss, collapse = "\n   ")))
cat("missing indicators: none\n")
cat(sprintf("         of which CCDR-98 : %d economies\n", uniqueN(long[is_ccdr == TRUE]$iso3)))
if (length(UNMATCHED)) {
  cat("\nlabels dropped (aggregates and non-countries), by source:\n")
  for (k in names(UNMATCHED))
    cat(sprintf("   %-16s %3d  e.g. %s\n", k, length(UNMATCHED[[k]]),
                paste(head(UNMATCHED[[k]], 4), collapse = ", ")))
}
