// Build CPAT-AI-Mitigation-MVP_AdvancedFeatures_v1.1.pptx (v1.1: next step names the first-run test macro): 4 slides on what v1.01-v1.03 added (cap-based ETS,
// multiple scenarios with a macro, stored results and comparison), in the style of the v1.0 overview deck.
// Screenshots and chart data: presentation/img/1x_*.png and advanced_data.json from make_screenshots_advanced_v1_0.py.
//
//   PPTX_SKILL=<pptx skill dir> NODE_PATH=<node_modules with pptxgenjs, image-size> node presentation/build_deck_advanced_v1_1.js
const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { imageSize } = require("image-size");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const HERE = __dirname;
const IMG = (f) => path.join(HERE, "img", f);
const OUT = path.join(HERE, "CPAT-AI-Mitigation-MVP_AdvancedFeatures_v1.1.pptx");
const DATA = JSON.parse(fs.readFileSync(IMG("advanced_data.json"), "utf8"));

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
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in, as the v1.0 overview deck
pres.title = "CPAT-AI-Mitigation-MVP v1.03: advanced features";
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
    { text: { text: "CPAT-AI-Mitigation-MVP v1.03-v1.04  ·  advanced features  ·  screenshots and numbers from the workbook",
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
function fit(file, maxW, maxH) {           // picture size within maxW x maxH, aspect ratio kept
  const d = imageSize(fs.readFileSync(IMG(file)));
  const w = Math.min(maxW, maxH * d.width / d.height);
  return { w, h: w * d.height / d.width };
}
function picture(slide, file, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, { x: x - 0.06, y: y - 0.06, w: w + 0.12, h: h + 0.12, fill: { color: "FFFFFF" },
    line: { color: LINE, width: 1 }, shadow: { type: "outer", color: "1E2761", opacity: 0.18, blur: 6, offset: 2, angle: 90 },
    objectName: nm("frame") });
  slide.addImage({ path: IMG(file), x, y, w, h, objectName: nm("screenshot") });
}
function badge(slide, txt, x, y, color, size = 0.42) {
  slide.addShape(pres.shapes.OVAL, { x, y, w: size, h: size, fill: { color }, line: { color }, objectName: nm("badge") });
  slide.addText(txt, { x, y, w: size, h: size, align: "center", valign: "middle", fontSize: 14, bold: true,
    color: "FFFFFF", margin: 0, isTextBox: true, objectName: nm("badgeText") });
}
function item(slide, num, head, body, x, y, w, color = TEAL, h = 0.95) {
  badge(slide, String(num), x, y, color);
  slide.addText([
    { text: head, options: { bold: true, fontSize: 15, color: INK, breakLine: true } },
    { text: body, options: { fontSize: 13, color: GREY } },
  ], { x: x + 0.6, y: y - 0.05, w: w - 0.6, h, valign: "top", margin: 0, isTextBox: true, objectName: nm("item") });
}
function caption(slide, text, x, y, w) {
  slide.addText(text, { x, y, w, h: 0.3, fontSize: 12, italic: true, color: GREY, margin: 0, isTextBox: true,
    objectName: nm("caption") });
}
const AXIS = { catAxisLabelColor: GREY, valAxisLabelColor: GREY, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", titleFontFace: "+mn-lt", titleFontSize: 13,
  titleColor: INK, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 11, dataLabelColor: INK };

// 1. Title and what is new ----------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "TITLE_DARK" });
  s.addText("ADVANCED FEATURES  ·  v1.01 TO v1.03", { x: 0.6, y: 1.05, w: 6, h: 0.35, fontSize: 12, bold: true,
    color: GOLD, charSpacing: 3, margin: 0, isTextBox: true, objectName: nm("kicker") });
  s.addText("CPAT-AI-Mitigation-MVP v1.03", { placeholder: "title" });
  s.addText("Since v1.00 the MVP prices an ETS from its cap, runs many scenarios in one go, and stores and compares "
    + "their results. AI wrote all of it, including the Excel macro.", { x: 0.6, y: 3.25, w: 5.9, h: 1.3,
    fontSize: 18, color: ICE, valign: "top", margin: 0, isTextBox: true, objectName: nm("subtitle") });
  ["Cap-based ETS", "Batch scenarios", "Stored results"].forEach((t, i) => {
    const w = [1.95, 2.0, 1.85][i], x = [0.6, 2.7, 4.85][i];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.85, w, h: 0.45, rectRadius: 0.12, fill: { color: "2E3A7A" },
      line: { color: "3A4A8A" }, objectName: nm("chip") });
    s.addText(t, { x, y: 4.85, w, h: 0.45, align: "center", valign: "middle", fontSize: 13, color: "FFFFFF",
      margin: 0, isTextBox: true, objectName: nm("chipText") });
  });
  s.addText("Draft, October 2026  ·  Egypt data  ·  for the team", { x: 0.6, y: 6.8, w: 6, h: 0.3, fontSize: 10,
    color: ICE, margin: 0, isTextBox: true, objectName: nm("footnote") });
  const rel = [
    ["v1.01", "Results in one place", "MTOutputs collects 25 key results per scenario by output code; Charts plots them."],
    ["v1.02", "A real ETS", "Cap against baseline emissions, fast price estimate, goal seek, benchmarks and a "
      + "volatility adjustment."],
    ["v1.03", "Many scenarios", "Scenario table in MTInputs, a macro that runs it, results stored as values, and a "
      + "comparison tab."],
  ];
  rel.forEach(([v, h, b], i) => {
    const y = 1.2 + i * 1.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.3, y, w: 5.4, h: 1.5, rectRadius: 0.1, fill: { color: "FFFFFF" },
      line: { color: "FFFFFF" }, objectName: nm("release") });
    s.addShape(pres.shapes.OVAL, { x: 7.55, y: y + 0.3, w: 0.9, h: 0.9, fill: { color: [TEAL, AMBER, NAVY][i] },
      line: { color: [TEAL, AMBER, NAVY][i] }, objectName: nm("releaseBadge") });
    s.addText(v, { x: 7.55, y: y + 0.3, w: 0.9, h: 0.9, align: "center", valign: "middle", fontSize: 14, bold: true,
      color: "FFFFFF", margin: 0, isTextBox: true, objectName: nm("releaseNo") });
    s.addText([
      { text: h, options: { bold: true, fontSize: 16, color: NAVY, breakLine: true } },
      { text: b, options: { fontSize: 13, color: INK } },
    ], { x: 8.65, y: y + 0.15, w: 3.85, h: 1.2, valign: "middle", margin: 0, isTextBox: true, objectName: nm("releaseText") });
  });
  s.addNotes("Three releases since the v1.00 overview deck: results collected by code (v1.01), an ETS that finds its "
    + "own price from a cap (v1.02), and multiple scenarios with a macro, stored results and a comparison tab (v1.03).");
}

// 2. New ETS ------------------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "NEW ETS", "The ETS now finds its own price from a cap");
  const steps = [
    ["1", "Set a cap", "A change against baseline covered emissions, stored as data from scenario 1 (no circular "
      + "reference)."],
    ["2", "Estimate the price", "LN(cap ÷ baseline) ÷ (semi-elasticity × effectiveness × volatility adjustment)."],
    ["3", "Goal seek", "Legacy's damped iteration until covered emissions are within 0.5% of the cap, as a macro "
      + "or in Python."],
  ];
  steps.forEach(([k, h, b], i) => {
    const x = 0.6 + i * 4.15;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.5, w: 3.75, h: 1.25, fill: { color: CARD }, line: { color: LINE },
      objectName: nm("stepCard") });
    item(s, k, h, b, x + 0.15, 1.65, 3.5, TEAL, 1.05);
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 1.98, w: 0.3, h: 0.3, fill: { color: GOLD },
      line: { color: GOLD }, objectName: nm("stepArrow") });
  });
  // split of the permit price
  s.addText("THE PRICE SPLITS IN TWO", { x: 0.6, y: 3.2, w: 6.5, h: 0.3, fontSize: 11, bold: true, color: TEAL,
    charSpacing: 2, margin: 0, isTextBox: true, objectName: nm("splitHead") });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.3, w: 1.75, h: 1.2, rectRadius: 0.1, fill: { color: NAVY },
    line: { color: NAVY }, objectName: nm("priceBox") });
  s.addText([
    { text: "Permit price", options: { bold: true, fontSize: 15, breakLine: true } },
    { text: "× volatility adjustment", options: { fontSize: 12 } },
  ], { x: 0.65, y: 4.3, w: 1.65, h: 1.2, align: "center", valign: "middle", color: "FFFFFF", margin: 0,
    isTextBox: true, objectName: nm("priceText") });
  const parts = [
    ["Auctioned share  × (1 − benchmark)", "Enters fuel prices like a carbon tax: usage and efficiency respond; "
      + "raises revenue.", TEAL_T, "C7E3DE", TEAL],
    ["Free allocation  × benchmark", "A shadow price on efficiency only, like feebates; no change in prices or "
      + "revenue.", AMBER_T, "F2D9B6", AMBER],
  ];
  parts.forEach(([h, b, fill, line, col], i) => {
    const y = 3.65 + i * 1.4;
    s.addShape(pres.shapes.RIGHT_ARROW, { x: 2.5, y: y + 0.42, w: 0.45, h: 0.32, fill: { color: GOLD }, line: { color: GOLD },
      objectName: nm("splitArrow") });
    s.addShape(pres.shapes.RECTANGLE, { x: 3.1, y, w: 4.0, h: 1.15, fill: { color: fill }, line: { color: line },
      objectName: nm("partCard") });
    s.addText([
      { text: h, options: { bold: true, fontSize: 14, color: col, breakLine: true } },
      { text: b, options: { fontSize: 13, color: INK } },
    ], { x: 3.25, y: y + 0.1, w: 3.75, h: 1.0, valign: "top", margin: 0, isTextBox: true, objectName: nm("partText") });
  });
  s.addText("Benchmarks by sector group replace the single auction share. Volatility adjustment = 1/1.1 at Medium, as "
    + "legacy.", { x: 0.6, y: 6.5, w: 6.5, h: 0.4, fontSize: 12, italic: true, color: GREY, margin: 0, isTextBox: true,
    objectName: nm("splitNote") });
  // chart: package B price path
  const yrs = DATA.years, i0 = yrs.indexOf(2027);
  s.addChart(pres.charts.LINE, [{ name: "ETS price", labels: yrs.slice(i0).map((y) => ([2027, 2030, 2035, 2040].includes(y) ? String(y) : "")),
    values: DATA.ets_p.slice(i0).map((v) => Math.round(v * 10) / 10) }], {
    x: 7.55, y: 3.15, w: 5.15, h: 3.3, showTitle: true, title: "Package B: ETS price after goal seek ($/tCO2)",
    chartColors: [TEAL], lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6, showLegend: false,
    valGridLine: { color: "E6E9EF", size: 0.75 }, catGridLine: { style: "none" }, ...AXIS,
    valAxisMinVal: 0, objectName: nm("etsChart") });
  s.addText("Covered emissions stay within 0.5% of the cap (−5% in 2027 to −20% by 2035) in every year.",
    { x: 7.55, y: 6.5, w: 5.15, h: 0.4, fontSize: 12, italic: true, color: GREY, margin: 0, isTextBox: true,
      objectName: nm("chartNote") });
  s.addNotes("The cap is set against baseline covered emissions, which are stored as data from scenario 1 to avoid "
    + "a circular reference. A fast closed-form estimate gives the starting price; the goal seek (legacy's damped "
    + "log-space iteration) then meets the cap. The permit price splits by the benchmarks: the auctioned part acts like "
    + "a carbon tax, the free part like a feebate. Package B prices: about $19 in 2027, $41 in 2030, $77 in 2035.");
}

// 3. Scenario table and macro -------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "MULTIPLE SCENARIOS · SET UP AND RUN", "One table of scenarios, one macro to run them");
  const p = fit("12_definitions.png", 7.3, 4.0);
  picture(s, "12_definitions.png", 0.6, 1.55, p.w, p.h);
  caption(s, "MTInputs: J baseline, K live scenario, L onwards one column per scenario (row 4: Run? = Yes)", 0.6,
    1.55 + p.h + 0.15, 7.3);
  s.addText("RunAllScenarios  (Alt+F8)", { x: 8.35, y: 1.5, w: 4.35, h: 0.35, fontSize: 15, bold: true, color: NAVY,
    margin: 0, isTextBox: true, objectName: nm("macroHead") });
  const steps = [
    ["Stores the baseline once", "It does not change unless you ask."],
    ["Copies each scenario into K", "Every column marked Run? = Yes, one at a time."],
    ["Recalculates", "And runs the ETS goal seek when the scenario has an ETS."],
    ["Pastes results as values", "Into StoredResults, under the scenario's number."],
    ["Puts K back", "The live scenario is as you left it."],
  ];
  steps.forEach(([h, b], i) => item(s, i + 1, h, b, 8.35, 2.05 + i * 0.95, 4.35, TEAL, 0.85));
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 5.45, w: 7.3, h: 1.0, fill: { color: CARD }, line: { color: LINE },
    objectName: nm("card") });
  s.addText([
    { text: "Shipped packages:  ", options: { bold: true, color: NAVY } },
    { text: "A  carbon tax to $50 by 2030 + feebates;  B  ETS on power and industry + $25 carbon tax elsewhere.",
      options: { color: INK, breakLine: true } },
    { text: "Other macros:  ", options: { bold: true, color: NAVY } },
    { text: "StoreBaseline, StoreLiveScenario, SolveETSLive, ClearStoredScenarios.", options: { color: INK } },
  ], { x: 0.75, y: 5.5, w: 7.0, h: 0.9, fontSize: 13, valign: "middle", margin: 0, isTextBox: true,
    objectName: nm("cardText") });
  s.addNotes("Scenario definitions follow legacy CPAT: one MTInputs column per scenario. The macro copies each one "
    + "marked Yes into the live column, recalculates, runs the ETS goal seek where needed and stores the MTOutputs "
    + "block as values. To add a scenario, copy the last definition column to the right and edit it.");
}

// 4. Compare and quality ------------------------------------------------------------------------------------------
{
  const s = pres.addSlide({ masterName: "CONTENT" });
  header(s, "MULTIPLE SCENARIOS · COMPARE", "Stored results, compared in one view");
  const p = fit("11_compare.png", 7.3, 3.0);
  picture(s, "11_compare.png", 0.6, 1.55, p.w, p.h);
  caption(s, "ScenarioCompare: choose a year and the scenario IDs; levels, differences and % vs the baseline",
    0.6, 1.55 + p.h + 0.15, 7.3);
  const ids = ["1", "3", "4", "5"], lab = ["Baseline", "$20 carbon price", "Package A", "Package B"];
  s.addChart(pres.charts.BAR, [{ name: "CO2 2030", labels: lab,
    values: ids.map((k) => Math.round(DATA.co2_2030[k] * 10) / 10) }], {
    x: 8.3, y: 1.45, w: 4.4, h: 3.35, barDir: "bar", showTitle: true, title: "CO2 from fuel combustion, 2030 (Mt)",
    chartColors: [NAVY], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0", showLegend: false,
    valGridLine: { color: "E6E9EF", size: 0.75 }, catGridLine: { style: "none" }, valAxisHidden: true,
    catAxisOrientation: "maxMin", barGapWidthPct: 60, ...AXIS, objectName: nm("co2Chart") });
  const stats = [
    ["8e-15", "Largest relative difference between the macro run in LibreOffice and the Python emulation"],
    ["0.5%", "Largest gap between the ETS package's covered emissions and its cap"],
    ["0", "Change in the 3,642 output codes shared with v1.02"],
  ];
  stats.forEach(([v, t], i) => {
    const x = 0.6 + i * 2.5;
    s.addText(v, { x, y: 5.3, w: 2.3, h: 0.65, fontSize: 34, bold: true, color: TEAL, fontFace: "Cambria", margin: 0,
      isTextBox: true, objectName: nm("stat") });
    s.addText(t, { x, y: 5.95, w: 2.3, h: 0.85, fontSize: 12, color: GREY, valign: "top", margin: 0, isTextBox: true,
      objectName: nm("statText") });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.3, y: 5.05, w: 4.4, h: 1.75, fill: { color: NAVY }, line: { color: NAVY },
    objectName: nm("nextCard") });
  s.addText([
    { text: "NEXT", options: { bold: true, fontSize: 11, color: GOLD, charSpacing: 2, breakLine: true } },
    { text: "Run CheckBatchRun once in Excel (one-click test)", options: { bullet: true, breakLine: true } },
    { text: "Power sector, CH4 and N2O", options: { bullet: true, breakLine: true } },
    { text: "Partial adjustment of fuel use; review benchmarks", options: { bullet: true } },
  ], { x: 8.5, y: 5.15, w: 4.05, h: 1.55, fontSize: 13, color: "FFFFFF", valign: "top", margin: 0, paraSpaceAfter: 4,
    isTextBox: true, objectName: nm("nextText") });
  s.addNotes("ScenarioCompare reads only the stored values, so results from earlier runs or other files can be "
    + "compared too. AI wrote the VBA project from the file-format specification, without Excel; LibreOffice runs it "
    + "and reproduces the Python emulation exactly. Its first run in Excel is one click: the macro CheckBatchRun checks "
    + "the lookups, compares Excel's recalculation with the stored results and reports PASS on sheet MacroCheck.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
