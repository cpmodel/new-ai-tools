// Build CPAT-AI-Mitigation-MVP_v2.3.pptx (v2.3: workbook v2.3; earlier decks in Old/): the combined, simple deck of
// the whole MVP (replaces the Overview v1.0 and
// Advanced features v1.1 decks). Same style: navy / teal / amber, Cambria headings, Calibri body.
// Screenshots and chart data: presentation/img/v23_*.png and img/v23_data.json from make_screenshots_v2_3.py
// (rendered from CPAT-AI-Mitigation-MVP-v2.3.xlsm). Numbers from check_report_v2.3.md (FACTS below).
//
//   node presentation/build_deck_v2_3.js
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { imageSize } = require("image-size");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const HERE = __dirname;
const IMG = (f) => path.join(HERE, "img", f);
const OUT = path.join(HERE, "CPAT-AI-Mitigation-MVP_v2.3.pptx");
const DATA = JSON.parse(fs.readFileSync(IMG("v23_data.json"), "utf8"));
const FACTS = JSON.parse(fs.readFileSync(IMG("v23_facts.json"), "utf8"));

const THEME = {
  name: "CPAT AI Mitigation",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1F1F2E", lt1: "FFFFFF", dk2: "1E2761", lt2: "F4F6FB",
    accent1: "1E2761", accent2: "0F7B6C", accent3: "E08A1E", accent4: "D4A24C",
    accent5: "5C667A", accent6: "D5DAE3", hlink: "0F7B6C", folHlink: "5C667A",
  },
};
const NAVY = "1E2761", TEAL = "0F7B6C", AMBER = "E08A1E", GOLD = "D4A24C", INK = "1F1F2E", GREY = "5C667A";
const CARD = "F4F6FB", LINE = "D5DAE3", TEAL_T = "EAF4F2", AMBER_T = "FDF3E6", ICE = "CADCFC";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.title = "CPAT-AI-Mitigation-MVP: overview and user guide";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

pres.defineSlideMaster({
  title: "TITLE_DARK",
  background: { color: NAVY },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.55, w: 6.2, h: 1.6, fontFace: "Cambria",
      fontSize: 40, bold: true, color: "FFFFFF", valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { options: { name: "label", type: "body", x: 0.6, y: 0.32, w: 12.1, h: 0.3, fontFace: "Calibri",
      fontSize: 11, bold: true, color: TEAL, charSpacing: 2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.62, w: 12.1, h: 0.7, fontFace: "Cambria",
      fontSize: 30, bold: true, color: NAVY, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "CPAT-AI-Mitigation-MVP v2.3  ·  overview and user guide  ·  screenshots and charts from the workbook",
      options: { x: 0.6, y: 7.05, w: 9.5, h: 0.25, fontSize: 9, color: GREY, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.25, fontSize: 9, color: GREY, align: "right" },
});

let n = 0;
const nm = (s) => `${s}-${++n}`;
function header(slide, label, title) {
  slide.addText(label, { placeholder: "label" });
  slide.addText(title, { placeholder: "title" });
}
// picture fitted into a box (aspect ratio kept), centred horizontally in the box; returns the size used
function picture(slide, file, x, y, maxW, maxH, align = "center") {
  const d = imageSize(fs.readFileSync(IMG(file)));
  let w = maxW, h = maxW * d.height / d.width;
  if (h > maxH) { h = maxH; w = maxH * d.width / d.height; }
  const px = align === "center" ? x + (maxW - w) / 2 : x;
  slide.addShape(pres.shapes.RECTANGLE, { x: px - 0.06, y: y - 0.06, w: w + 0.12, h: h + 0.12, fill: { color: "FFFFFF" },
    line: { color: LINE, width: 1 }, shadow: { type: "outer", color: "1E2761", opacity: 0.18, blur: 6, offset: 2, angle: 90 },
    objectName: nm("frame") });
  slide.addImage({ path: IMG(file), x: px, y, w, h, objectName: nm("screenshot") });
  return { x: px, y, w, h };
}
function caption(slide, text, x, y, w) {
  slide.addText(text, { x, y, w, h: 0.3, fontSize: 12, italic: true, color: GREY, margin: 0, isTextBox: true,
    objectName: nm("caption") });
}
function badge(slide, txt, x, y, color, size = 0.42) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: size, h: size, fill: { color }, line: { color }, objectName: nm("badge") });
  slide.addText(txt, { x, y, w: size, h: size, align: "center", valign: "middle", fontSize: 14, bold: true,
    color: "FFFFFF", margin: 0, isTextBox: true, objectName: nm("badgeText") });
}
function item(slide, num, head, body, x, y, w, color = TEAL, h = 1.0) {
  badge(slide, String(num), x, y, color);
  slide.addText([
    { text: head, options: { bold: true, fontSize: 15, color: INK, breakLine: true } },
    { text: body, options: { fontSize: 14, color: GREY } },
  ], { x: x + 0.6, y: y - 0.05, w: w - 0.6, h, valign: "top", margin: 0, isTextBox: true, objectName: nm("item") });
}
function card(slide, x, y, w, h, fill, lineC, runs) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: lineC }, objectName: nm("card") });
  slide.addText(runs, { x: x + 0.18, y: y + 0.12, w: w - 0.36, h: h - 0.24, valign: "top", margin: 0, isTextBox: true,
    objectName: nm("cardText") });
}
function stat(slide, big, small, x, y, w, color = NAVY) {
  slide.addText([
    { text: big, options: { fontSize: 40, bold: true, color, fontFace: "Cambria", breakLine: true } },
    { text: small, options: { fontSize: 13, color: GREY } },
  ], { x, y, w, h: 1.5, valign: "top", margin: 0, isTextBox: true, objectName: nm("stat") });
}
function flow(slide, labels, y, color = NAVY) {
  const k = labels.length, gap = 0.4, w = (12.1 - gap * (k - 1)) / k;
  labels.forEach((t, i) => {
    const x = 0.6 + i * (w + gap);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h: 0.62, rectRadius: 0.1, fill: { color }, line: { color },
      objectName: nm("flow") });
    slide.addText(t, { x, y, w, h: 0.62, align: "center", valign: "middle", fontSize: 14, bold: true, color: "FFFFFF",
      margin: 0, isTextBox: true, objectName: nm("flowText") });
    if (i < k - 1) slide.addShape(pres.shapes.RIGHT_ARROW, { x: x + w + 0.07, y: y + 0.17, w: gap - 0.14, h: 0.28,
      fill: { color: GOLD }, line: { color: GOLD }, objectName: nm("flowArrow") });
  });
}
const CHART_TXT = { catAxisLabelColor: GREY, valAxisLabelColor: GREY, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", legendFontFace: "+mn-lt" };
function lineChart(slide, series, x, y, w, h, title, colors, valFmt = "#,##0") {
  slide.addChart(pres.charts.LINE, series, {
    x, y, w, h, showTitle: true, title, titleFontSize: 14, titleColor: INK, chartColors: colors, lineSize: 2.5,
    lineDataSymbol: "none", showLegend: series.length > 1, legendPos: "b", legendFontSize: 11, legendColor: GREY,
    valGridLine: { color: "E3E7EF", size: 0.75 }, catGridLine: { style: "none" }, valAxisLabelFormatCode: valFmt,
    catAxisLabelFrequency: 3, catAxisLabelRotate: 315, ...CHART_TXT, objectName: nm("chart"),
  });
}
const years = DATA.years.map(String);
const ids = ["1", "3", "4", "5"].filter((k) => DATA.co2[k]);
const short = { "1": "Baseline", "3": "Carbon price $20", "4": "Package A", "5": "Package B" };
const SCEN_COLORS = ["5C667A", "0F7B6C", "E08A1E", "1E2761"];

// 1. Title ------------------------------------------------------------------------------------------------------
pres.addSection({ title: "Overview" });
{
  const s = pres.addSlide({ masterName: "TITLE_DARK", sectionTitle: "Overview" });
  s.addText("OVERVIEW + USER GUIDE", { x: 0.6, y: 1.05, w: 6, h: 0.35, fontSize: 12, bold: true, color: GOLD,
    charSpacing: 3, margin: 0, isTextBox: true, objectName: nm("kicker") });
  s.addText("CPAT-AI-Mitigation-MVP", { placeholder: "title" });
  s.addText("The CPAT mitigation module as a short, readable Excel workbook, generated and checked by AI.",
    { x: 0.6, y: 3.25, w: 5.9, h: 1.2, fontSize: 18, color: ICE, valign: "top", margin: 0, isTextBox: true,
      objectName: nm("subtitle") });
  ["Plain formulas", "Scenarios side by side", "Checked by Python"].forEach((t, i) => {
    const w = [1.9, 2.55, 2.15][i], x = [0.6, 2.65, 5.35][i];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.75, w, h: 0.45, rectRadius: 0.12, fill: { color: "2E3A7A" },
      line: { color: "3A4A8A" }, objectName: nm("chip") });
    s.addText(t, { x, y: 4.75, w, h: 0.45, align: "center", valign: "middle", fontSize: 13, color: "FFFFFF",
      margin: 0, isTextBox: true, objectName: nm("chipText") });
  });
  s.addText("Version 2.3, October 2026  ·  Egypt data  ·  prototype for discussion", { x: 0.6, y: 6.8, w: 6.5,
    h: 0.3, fontSize: 10, color: ICE, margin: 0, isTextBox: true, objectName: nm("footnote") });
  const p = picture(s, "v23_01_overview.png", 7.9, 1.2, 4.8, 4.2);
  s.addText("The Mitigation tab as it opens: everything rolled up", { x: p.x, y: p.y + p.h + 0.15, w: 4.8, h: 0.3,
    fontSize: 11, color: ICE, italic: true, margin: 0, isTextBox: true, objectName: nm("caption") });
  s.addNotes("One workbook for the CPAT mitigation chain: policies, prices, power, fuel use, revenues and CO2. "
    + "Generated and checked by AI; every value is recomputed independently.");
}

// 2. What it covers --------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Overview" });
  header(s, "WHAT IT IS", "One tab follows the CPAT mitigation chain");
  flow(s, ["1. Policies", "2. Prices", "3. Power", "5-8. Fuel use", "11-12. Energy, revenue", "13. CO2"], 1.65);
  item(s, 1, "Policies", "Carbon tax, ETS with a cap, feebates, fuel price reform; existing carbon pricing.", 0.6, 2.75, 6);
  item(s, 2, "Prices and power", "Retail prices by fuel; a power model with demand, dispatch, retirement and "
    + "investment.", 0.6, 3.85, 6);
  item(s, 3, "Fuel use and results", "16 sectors × 8 fuels plus power; fuel use, revenues, CO2 to 2041, real 2026 USD.",
    0.6, 4.95, 6);
  card(s, 7.1, 2.7, 5.6, 2.75, CARD, LINE, [
    { text: "IN THE WORKBOOK", options: { bold: true, fontSize: 11, color: TEAL, charSpacing: 2, breakLine: true } },
    { text: "Egypt data, 2022-2041", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Any number of scenarios, run by one macro", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Results and charts by output code", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "NOT YET", options: { bold: true, fontSize: 11, color: AMBER, charSpacing: 2, breakLine: true } },
    { text: "Other gases and local pollutants", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Full numerical match with legacy CPAT", options: { bullet: true, fontSize: 15, color: INK } },
  ]);
  s.addNotes("The section numbers follow legacy CPAT. Power demand has its own elasticities; sector policies act on "
    + "fuels only (electricity appears as memo lines in the sector blocks).");
}

// 3. Problems and fixes ------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Overview" });
  header(s, "WHAT CHANGES", "Four problems with today's Excel, four fixes");
  s.addText("TODAY IN CPAT", { x: 0.6, y: 1.5, w: 5.2, h: 0.3, fontSize: 11, bold: true, color: AMBER, charSpacing: 2,
    margin: 0, isTextBox: true, objectName: nm("colhead") });
  s.addText("IN THE MVP", { x: 6.7, y: 1.5, w: 6, h: 0.3, fontSize: 11, bold: true, color: TEAL, charSpacing: 2,
    margin: 0, isTextBox: true, objectName: nm("colhead") });
  const rows = [
    ["Hard to read: long formulas, many tabs", "One tab, numbered sections; every formula at most 100 characters"],
    ["Formulas differ from row to row", "One formula per block: copy, drag or paste it anywhere in the block"],
    ["New scenarios are slow to build", "Copy a scenario block, or list scenarios in MTInputs and run one macro"],
    ["Results are hard to verify", "Every value recomputed in Python; legacy bugs found and listed"],
  ];
  rows.forEach(([a, b], i) => {
    const y = 1.9 + i * 1.22;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y, w: 5.2, h: 1.02, fill: { color: AMBER_T }, line: { color: "F2D9B6" },
      objectName: nm("problem") });
    s.addText(a, { x: 0.8, y, w: 4.85, h: 1.02, fontSize: 16, color: INK, valign: "middle", margin: 0, isTextBox: true,
      objectName: nm("problemText") });
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 5.95, y: y + 0.33, w: 0.6, h: 0.36, fill: { color: GOLD }, line: { color: GOLD },
      objectName: nm("arrow") });
    s.addShape(pres.shapes.RECTANGLE, { x: 6.7, y, w: 6.0, h: 1.02, fill: { color: TEAL_T }, line: { color: "C7E3DE" },
      objectName: nm("fix") });
    s.addText(b, { x: 6.9, y, w: 5.65, h: 1.02, fontSize: 16, color: INK, valign: "middle", margin: 0, isTextBox: true,
      objectName: nm("fixText") });
  });
  s.addNotes("Readability, copy-pasteable formulas, fast scenarios, and independent verification.");
}

// 4. Scenarios side by side ------------------------------------------------------------------------------------
pres.addSection({ title: "Using the workbook" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using the workbook" });
  header(s, "USER GUIDE · LAYOUT", "Years run across, scenarios sit side by side");
  const p = picture(s, "v23_01b_scenarios.png", 0.6, 1.55, 12.1, 2.2);
  caption(s, "Each scenario is a block of columns: output codes, 2022 (beige), the years to 2041, a blank column.",
    0.6, p.y + p.h + 0.15, 12.1);
  item(s, 1, "Copy a settings column", "In MTInputs, copy the last scenario column one to the right and edit it.",
    0.6, p.y + p.h + 0.8, 4.0);
  item(s, 2, "Copy a scenario block", "In Mitigation, copy a whole block and paste it to the right.", 4.65,
    p.y + p.h + 0.8, 4.0);
  item(s, 3, "Done", "The scenario number goes up by one; the block reads its own MTInputs column.", 8.7,
    p.y + p.h + 0.8, 4.0);
  s.addNotes("Row 5 holds the scenario number; every input row looks up its value in that MTInputs column.");
}

// 5. Left section ------------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using the workbook" });
  header(s, "USER GUIDE · ROWS", "Row-specific numbers live on the left");
  picture(s, "v23_01c_left.png", 0.6, 1.55, 7.9, 4.9, "left");
  item(s, "A", "Columns B:C: fuel and sector", "Every row is one fuel in one sector (here: road transport).",
    9.0, 1.65, 3.75, NAVY, 1.2);
  item(s, "B", "Columns D:G: that row's numbers", "Emission factor, elasticities, positions: looked up once.",
    9.0, 3.05, 3.75, NAVY, 1.2);
  item(s, "C", "Year columns: one formula", "The same formula in every row of a block.", 9.0, 4.45, 3.75, NAVY, 1.2);
  s.addNotes("Columns B:G and I:K are grouped and rolled up on open: click + above the columns to show them. "
    + "The italic memo line under each sector heading is its electricity demand from the power model.");
}

// 6. Formulas ----------------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using the workbook" });
  header(s, "USER GUIDE · FORMULAS", "Short, plain formulas; the last year runs on named LAMBDAs");
  const p = picture(s, "v23_02_formulas.png", 0.6, 1.5, 12.1, 2.6);
  const y0 = p.y + p.h + 0.45, hc = 1.35;
  card(s, 0.6, y0, 3.85, hc, CARD, LINE, [
    { text: "At most 100 characters", options: { bold: true, fontSize: 15, color: NAVY, breakLine: true } },
    { text: `All ${FACTS.formulas} formulas. Long ones are split into visible steps.`,
      options: { fontSize: 14, color: INK } }]);
  card(s, 4.72, y0, 3.85, hc, TEAL_T, "C7E3DE", [
    { text: "2022-2040: no LAMBDA needed", options: { bold: true, fontSize: 15, color: TEAL, breakLine: true } },
    { text: "Plain Excel from 2019; LibreOffice recalculates these years without errors.",
      options: { fontSize: 14, color: INK } }]);
  card(s, 8.85, y0, 3.85, hc, "EEECE1", "D9D4BC", [
    { text: "2041: the LAMBDA year", options: { bold: true, fontSize: 15, color: "9C0006", breakLine: true } },
    { text: "Every cell of 2041 calls a named LAMBDA (Excel 365); sheet Lambdas lists them.",
      options: { fontSize: 14, color: INK } }]);
  s.addNotes("Formula text of the road gasoline rows, 2040 column: the same formula as every year to 2040; 2041 calls the named LAMBDAs. Lookups use "
    + "defined names for the data tables (e.g. WEOVal, PDomVal).");
}

// 7. Many scenarios ----------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Using the workbook" });
  header(s, "USER GUIDE · MANY SCENARIOS", "List scenarios, run one macro, compare");
  flow(s, ["Define in MTInputs", "RunAllScenarios", "StoredResults", "ScenarioCompare"], 1.6, TEAL);
  const p = picture(s, "v23_06_definitions.png", 0.6, 2.6, 6.0, 3.6, "left");
  caption(s, "MTInputs: one column per scenario, Run? = Yes", 0.6, p.y + p.h + 0.12, 6);
  const q = picture(s, "v23_05_compare.png", 6.9, 2.6, 5.8, 3.6, "left");
  caption(s, "ScenarioCompare: levels in one year vs the baseline", 6.9, q.y + q.h + 0.12, 5.8);
  s.addNotes("The macro copies each definition into the live column, recalculates, runs the ETS goal seek where "
    + "needed, and stores the results as values. CheckBatchRun tests the macros on first use in Excel.");
}

// 8. Policies and results ---------------------------------------------------------------------------------------
pres.addSection({ title: "Model and results" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Model and results" });
  header(s, "RESULTS", "Policies cut CO2; the ETS price meets its cap");
  lineChart(s, ids.map((k) => ({ name: short[k] || DATA.names[k], labels: years, values: DATA.co2[k] })),
    0.6, 1.5, 6.6, 4.9, "CO2 from fuel combustion incl. power (MtCO2)", SCEN_COLORS);
  lineChart(s, [{ name: "Covered CO2", labels: years, values: DATA.co2_ets },
                { name: "Cap", labels: years, values: DATA.co2_cap }],
    7.5, 1.5, 5.2, 4.9, "Package B: covered CO2 and cap (MtCO2)", [TEAL, AMBER]);
  caption(s, `Package B ETS price: ${FACTS.ets2030} $/tCO2 in 2030 (goal seek, ${FACTS.goalseek_iter} iterations)`,
    7.5, 6.5, 5.2);
  s.addNotes("Stored scenario results (StoredResults). Baseline = no carbon price; Package A = carbon tax to $50 + "
    + "feebates; Package B = ETS on power and industry + carbon tax elsewhere.");
}

// 9. Power -------------------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Model and results" });
  header(s, "POWER SECTOR", "A full power model, with ten legacy bugs fixed");
  flow(s, ["H. Demand", "I. Dispatch", "K. Capacity, retirement", "L. Investment", "G. Prices"], 1.6, NAVY);
  lineChart(s, ids.map((k) => ({ name: short[k] || DATA.names[k], labels: years, values: DATA.vre[k] })),
    0.6, 2.55, 6.3, 4.1, "Variable renewable share of generation", SCEN_COLORS, "0%");
  card(s, 7.2, 2.6, 5.5, 2.75, CARD, LINE, [
    { text: "SAME METHOD AS LEGACY", options: { bold: true, fontSize: 11, color: TEAL, charSpacing: 2, breakLine: true } },
    { text: "Reproduces 208 legacy rows of two runs (to 1e-13)", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Ten bugs fixed; report in Word", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: " ", options: { fontSize: 8, breakLine: true } },
    { text: "IN THE WORKBOOK", options: { bold: true, fontSize: 11, color: TEAL, charSpacing: 2, breakLine: true } },
    { text: "Planned retirement by type (switch per type)", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Electricity demand as memo lines by sector", options: { bullet: true, fontSize: 15, color: INK, breakLine: true } },
    { text: "Power fuel use, revenue and CO2 in the totals", options: { bullet: true, fontSize: 15, color: INK } },
  ]);
  s.addNotes("Demand responds to last year's price; dispatch fills demand from capacity; investment follows a logit on "
    + "levelised cost within scale-up limits. A carbon price raises the VRE share.");
}

// 10. Quality ----------------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Model and results" });
  header(s, "QUALITY", "Every number is checked independently");
  stat(s, "0", "error values when LibreOffice recalculates the shipped file", 0.6, 1.7, 2.8, TEAL);
  stat(s, FACTS.formulas, "formulas, none longer than 100 characters", 3.7, 1.7, 2.8);
  stat(s, FACTS.codes, "output codes unchanged against the previous version", 6.8, 1.7, 2.8);
  stat(s, FACTS.tests, "scenario tests recomputed in Python", 9.9, 1.7, 2.8, AMBER);
  item(s, 1, "Independent recomputation", "A Python script recomputes prices, fuel use, power, revenue and CO2 from "
    + "the source data for every test scenario.", 0.6, 3.9, 6.0);
  item(s, 2, "Macros tested", "The batch macro runs in LibreOffice and matches the stored results; CheckBatchRun "
    + "repeats the test in Excel.", 6.7, 3.9, 6.0);
  item(s, 3, "Differences listed", "Sheet LegacyDiff lists every difference with legacy CPAT and why.", 0.6, 5.3, 6.0);
  item(s, 4, "Documented", "Method notes, a Word documentation, and one register of bugs, issues and upgrades.", 6.7, 5.3, 6.0);
  s.addNotes(`Check report check_report_v2.3.md: Overall ${FACTS.overall}.`);
}

// 11. Getting started --------------------------------------------------------------------------------------------
pres.addSection({ title: "Next steps" });
{
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: "Next steps" });
  header(s, "GETTING STARTED", "Open, run the check, then explore");
  item(s, 1, "Open the .xlsm and enable macros", "Run CheckBatchRun once (Alt+F8): sheet MacroCheck should say PASS.",
    0.6, 1.7, 7.0);
  item(s, 2, "Change a scenario", "Edit its MTInputs column; results update on Mitigation, MTOutputs and Charts.",
    0.6, 2.9, 7.0);
  item(s, 3, "Run many scenarios", "Set Run? = Yes on the definitions and run RunAllScenarios; compare them.",
    0.6, 4.1, 7.0);
  item(s, 4, "Read more", "Quick start in the Word documentation; ReadMe and LegacyDiff tabs.", 0.6, 5.3, 7.0);
  s.addShape(pres.shapes.RECTANGLE, { x: 8.1, y: 1.5, w: 4.6, h: 5.0, fill: { color: NAVY }, line: { color: NAVY },
    objectName: nm("card") });
  s.addText([
    { text: "NEXT", options: { bold: true, fontSize: 11, color: GOLD, charSpacing: 2, breakLine: true } },
    { text: "Review in Excel", options: { bullet: true, breakLine: true } },
    { text: "Numerical check against legacy CPAT", options: { bullet: true, breakLine: true } },
    { text: "Other gases and local pollutants", options: { bullet: true, breakLine: true } },
    { text: "Planned-retirement data for Egypt", options: { bullet: true, breakLine: true } },
    { text: "Stress test: add a fuel, a sector, a scenario", options: { bullet: true } },
  ], { x: 8.35, y: 1.75, w: 4.1, h: 4.6, fontSize: 16, color: "FFFFFF", valign: "top", margin: 0, paraSpaceAfter: 14,
    isTextBox: true, objectName: nm("next") });
  s.addNotes("Files: CPAT-AI-Mitigation-MVP-v2.3.xlsm, CPAT-AI-Mitigation-MVP_Documentation_v2.3.docx, "
    + "CPAT-AI-Mitigation-MVP_BugsIssuesUpgrades_v2.3.docx (working_version folder).");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
