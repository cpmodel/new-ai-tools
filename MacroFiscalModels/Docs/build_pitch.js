// Builds Docs/MacroFiscalModels_Pitch_v0_1.pptx from Docs/deck_data.json.
// Needs pptxgenjs, react, react-dom, react-icons, sharp:  node build_pitch.js
// (python3 export_deck_data.py first refreshes the numbers.)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");

const HERE = __dirname;
const D = JSON.parse(fs.readFileSync(path.join(HERE, "deck_data.json"), "utf8"));
const OUT = path.join(HERE, "MacroFiscalModels_Pitch_v0_1.pptx");
const APPLY_THEME = process.env.APPLY_THEME || null;   // path to the pptx skill's apply_theme.js, if available

const HEX = { navy: "14213D", teal: "3D7A73", slate: "3F4A5A", grey: "5B6573", mid: "7A8494",
  orange: "A8521A", line: "E1E5EA", bg: "EAEEF2", white: "FFFFFF", sand: "E0A46E", lteal: "6FA8A0" };
const THEME = { name: "MacroFiscal", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: HEX.navy, lt1: HEX.white, dk2: HEX.slate, lt2: HEX.bg, accent1: HEX.teal, accent2: HEX.orange,
    accent3: HEX.navy, accent4: HEX.mid, accent5: HEX.lteal, accent6: HEX.sand, hlink: HEX.teal, folHlink: HEX.slate } };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";            // 13.33 x 7.5 in
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Macro-fiscal models for climate investment";
pres.author = "Stephen Stretton";
const C = pres.SchemeColor;
const FOOT = "Macro-fiscal models for climate investment";

// ---------------------------------------------------------------- layouts
pres.defineSlideMaster({ title: "DARK", background: { color: HEX.navy },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 1.9, w: 11.5, h: 0.4, fontSize: 12, bold: true,
      color: C.accent6, charSpacing: 2, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.35, w: 11.5, h: 1.7, fontSize: 40, bold: true, align: "left",
      color: C.background1, fontFace: "Cambria", valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.2, w: 11.0, h: 1.2, fontSize: 18,
      color: C.background2, valign: "top", margin: 0 }, text: "" } },
    { text: { text: "Stephen Stretton · World Bank", options: { x: 0.8, y: 6.7, w: 6, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ] });
pres.defineSlideMaster({ title: "CONTENT", background: { color: HEX.bg }, margin: [0.5, 0.6, 0.6, 0.6],
  slideNumber: { x: 12.3, y: 7.0, w: 0.5, h: 0.3, fontSize: 9, color: HEX.mid, align: "right" },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.4, w: 12, h: 0.3, fontSize: 11, bold: true,
      color: C.accent2, charSpacing: 2, valign: "top", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.72, w: 12.1, h: 0.75, fontSize: 30, bold: true, align: "left",
      color: C.text1, fontFace: "Cambria", valign: "top", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 9, color: C.accent4, margin: 0 } } },
  ] });

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function card(s, x, y, w, h, name) {
  s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.background1 }, line: { color: C.background1 },
    shadow: { type: "outer", color: "9AA3AE", blur: 4, offset: 1, angle: 90, opacity: 0.25 }, objectName: name });
}
function txt(s, text, o) { s.addText(text, Object.assign({ isTextBox: true, margin: 0, valign: "top", color: C.text2, fontSize: 14 }, o)); }
function content(sec, kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: sec });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function bullets(items, size = 14) {
  return items.map((t, i) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: i < items.length - 1, paraSpaceAfter: 6, fontSize: size } }));
}
const chartBase = (title) => ({
  showTitle: true, title, titleFontSize: 12, titleColor: HEX.slate, titleFontFace: "+mn-lt",
  catAxisLabelColor: HEX.grey, valAxisLabelColor: HEX.grey, catAxisLabelFontSize: 11, valAxisLabelFontSize: 11,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  valGridLine: { color: HEX.line, size: 0.75 }, catGridLine: { style: "none" },
  showLegend: true, legendPos: "b", legendFontSize: 11, legendColor: HEX.slate, legendFontFace: "+mn-lt",
  lineSize: 2.25, lineDataSymbol: "none", chartColors: [HEX.navy, HEX.orange, HEX.teal, HEX.mid],
});

(async () => {
  // 1 ------------------------------------------------------------ title
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Introduction" });
  s.addText("FOR DISCUSSION · OCTOBER 2026", { placeholder: "kicker" });
  s.addText("Macro-fiscal models for climate investment", { placeholder: "title" });
  s.addText("GIMF-lite, DIGNAD-lite and LIC-DSF-lite, linked to the national investment needs model", { placeholder: "body" });
  s.addNotes("Three simplified, open-source versions of IMF and World Bank tools, built so that a country's net zero investment programme from NMIM can be run through a macro model and a debt sustainability test.");

  // 2 ------------------------------------------------------------ the question
  s = content("Introduction", "THE QUESTION", "Investment needs tell us how much, not what it does");
  const steps = [
    [fa.FaCalculator, "1 · How much?", "NMIM", "USD bn a year by sector, 2025-2050, split by who pays: the budget, regulated utilities, private investors."],
    [fa.FaChartLine, "2 · What does it do?", "GIMF-lite · DIGNAD-lite", "Growth, consumption, inflation, the exchange rate, crowding out, and how the financing choice changes them."],
    [fa.FaBalanceScale, "3 · Is it affordable?", "LIC-DSF-lite", "Debt indicators against thresholds, stress tests and the risk of debt distress."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.85, 3.75, 3.85, `step${i + 1}`);
    s.addShape(pres.shapes.OVAL, { x: x + 0.35, y: 2.15, w: 0.8, h: 0.8, fill: { color: i === 2 ? C.accent2 : C.accent1 }, line: { color: C.background1 }, objectName: `step${i + 1}_circle` });
    s.addImage({ data: await icon(steps[i][0], HEX.white), x: x + 0.55, y: 2.35, w: 0.4, h: 0.4, altText: steps[i][1] });
    txt(s, steps[i][1], { x: x + 0.35, y: 3.15, w: 3.1, h: 0.45, fontSize: 20, bold: true, color: C.text1, fontFace: "Cambria" });
    txt(s, steps[i][2], { x: x + 0.35, y: 3.65, w: 3.1, h: 0.35, fontSize: 14, bold: true, color: C.accent2 });
    txt(s, steps[i][3], { x: x + 0.35, y: 4.1, w: 3.1, h: 1.5, fontSize: 14 });
    if (i < 2) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 3.8, y: 3.55, w: 0.3, h: 0.4, fill: { color: C.accent4 }, line: { color: C.accent4 }, objectName: `arrow${i + 1}` });
  }
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 6.0, w: 12.1, h: 0.6, fill: { color: C.text1 }, line: { color: C.text1 }, objectName: "so_box" });
  txt(s, [{ text: "So: ", options: { bold: true } }, { text: "one pipeline from investment needs to debt risk, country by country, in open code." }],
    { x: 0.85, y: 6.0, w: 11.7, h: 0.6, valign: "middle", color: C.background1, fontSize: 15 });
  s.addNotes("NMIM tells us the size and composition of investment. Ministries also ask what it does to growth and debt and whether it changes the risk of debt distress. The toolkit answers those three questions in one pipeline.");

  // 3 ------------------------------------------------------------ toolkit
  pres.addSection({ title: "The toolkit" });
  s = content("The toolkit", "THE TOOLKIT", "Three open-source versions of IMF and World Bank tools");
  const hdr = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 13 } });
  const cell = (t, b) => ({ text: t, options: { fontSize: 13, color: C.text2, bold: !!b, fill: { color: C.background1 } } });
  s.addTable([
    [hdr("Model"), hdr("Original"), hdr("Question here"), hdr("Kept"), hdr("Simplified")],
    [cell("GIMF-lite", 1), cell("IMF Global Integrated Monetary and Fiscal Model (Kumhof et al. 2010)"), cell("What does a net zero investment push do to the economy and public debt?"),
      cell("Non-Ricardian OLG and liquidity-constrained households; sticky prices and wages; Taylor rule; UIP; fiscal rule"), cell("One country, one good; no oil block or financial accelerator")],
    [cell("DIGNAD-lite", 1), cell("IMF Debt-Investment-Growth and Natural Disasters (Marto et al. 2018)"), cell("Public investment scale-ups and disasters in LICs: efficiency, financing, adaptation"),
      cell("Savers and hand-to-mouth; standard and adaptation capital; PIMA efficiency; grant / concessional / commercial debt"), cell("One good, fixed labour, no domestic debt")],
    [cell("LIC-DSF-lite", 1), cell("World Bank-IMF LIC Debt Sustainability Framework (2018)"), cell("Does the debt path change the risk of debt distress?"),
      cell("Composite indicator, thresholds, PV by loan cohort, five indicators, stress tests, rating"), cell("USD only; no realism tools or judgement")],
  ], { x: 0.6, y: 1.75, w: 12.1, colW: [1.5, 2.6, 2.7, 3.3, 2.0], border: { type: "solid", pt: 0.75, color: HEX.line }, rowH: [0.45, 1.25, 1.25, 1.1], valign: "middle", margin: 0.08 });
  txt(s, "Python, about 1,000 lines; a shared perfect-foresight solver; each model also runs on its own.", { x: 0.6, y: 6.25, w: 12, h: 0.4, fontSize: 13, italic: true, color: C.accent4 });
  s.addNotes("Each model keeps the mechanism that matters for the question and drops the rest. All calibrations are illustrative until a country is calibrated.");

  // 4 ------------------------------------------------------------ pipeline
  s = content("The toolkit", "HOW IT FITS TOGETHER", "One bridge feeds every model from the investment needs model");
  const box = (x, y, w, h, title, sub, fill, name, tcol) => {
    s.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, objectName: name });
    txt(s, title, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.4, fontSize: 16, bold: true, color: tcol || C.background1, fontFace: "Cambria" });
    txt(s, sub, { x: x + 0.2, y: y + 0.6, w: w - 0.4, h: h - 0.7, fontSize: 13, color: tcol || C.background1 });
  };
  const arrow = (x1, y1, x2, y2, n) => s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: C.accent4, width: 2, endArrowType: "triangle" }, objectName: n });
  box(0.6, 2.6, 2.4, 2.1, "NMIM v0.2", "186 countries, 2025-2050; investment by sector and by who pays", C.text1, "nmim");
  card(s, 3.6, 1.85, 2.7, 3.6, "bridge");
  txt(s, "nmim_bridge.py", { x: 3.8, y: 2.0, w: 2.4, h: 0.4, fontSize: 16, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, "% of GDP, each year:", { x: 3.8, y: 2.45, w: 2.4, h: 0.3, fontSize: 13, color: C.accent4 });
  txt(s, bullets(["Public budget need (P1 + public gap of P4)", "Private incremental investment", "Net energy savings"], 14), { x: 3.8, y: 2.85, w: 2.35, h: 2.4 });
  box(6.9, 1.85, 2.7, 1.6, "GIMF-lite", "All three paths; economy-wide effects", C.accent1, "gimf");
  box(6.9, 3.85, 2.7, 1.6, "DIGNAD-lite", "Public part; LIC financing and efficiency", C.accent1, "dignad");
  box(10.2, 2.6, 2.5, 2.1, "LIC-DSF-lite", "Growth and primary deficit shift the DSA baseline → rating", C.accent2, "dsf");
  arrow(3.0, 3.65, 3.6, 3.65, "a1"); arrow(6.3, 3.3, 6.9, 2.65, "a2"); arrow(6.3, 4.0, 6.9, 4.65, "a3");
  arrow(9.6, 2.65, 10.2, 3.3, "a4"); arrow(9.6, 4.65, 10.2, 4.0, "a5");
  txt(s, "python3 run_demo.py KEN  — any NMIM country, about 15 seconds", { x: 0.6, y: 6.0, w: 12, h: 0.4, fontSize: 14, fontFace: "Courier New", color: C.text1 });
  s.addNotes("The bridge runs the NMIM engine for one country and passes three paths to the macro models. The macro models' GDP and primary deficit paths then shift a LIC-DSF baseline.");

  // 5 ------------------------------------------------------------ GIMF mechanism
  pres.addSection({ title: "Results" });
  s = content("Results", "GIMF-LITE", "Financing matters because households are not Ricardian");
  const rows = [
    [fa.FaHourglassHalf, "20-year planning horizon", "Households discount future taxes at about (1 + r)/(θχ), above the government's rate, so debt feels like wealth."],
    [fa.FaWallet, "A quarter spend current income", "Liquidity-constrained households make transfers and temporary measures stronger."],
    [fa.FaCogs, "Sticky prices and wages, Taylor rule", "Output is demand-led in the short run; the exchange rate and current account adjust."],
  ];
  for (let i = 0; i < 3; i++) {
    const y = 1.85 + i * 1.45;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.75, h: 0.75, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `mech${i}` });
    s.addImage({ data: await icon(rows[i][0], HEX.white), x: 0.79, y: y + 0.19, w: 0.37, h: 0.37, altText: rows[i][1] });
    txt(s, rows[i][1], { x: 1.6, y, w: 5.6, h: 0.4, fontSize: 18, bold: true, color: C.text1, fontFace: "Cambria" });
    txt(s, rows[i][2], { x: 1.6, y: y + 0.45, w: 5.6, h: 0.85, fontSize: 14 });
  }
  card(s, 7.7, 1.85, 5.0, 4.4, "validation");
  txt(s, "VALIDATION", { x: 8.0, y: 2.1, w: 4.4, h: 0.3, fontSize: 11, bold: true, color: C.accent2, charSpacing: 2 });
  txt(s, "0.3 pp", { x: 8.0, y: 2.5, w: 4.4, h: 0.9, fontSize: 54, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, "gap in first-year consumption between debt- and tax-financed public investment (default calibration)", { x: 8.0, y: 3.45, w: 4.4, h: 0.8, fontSize: 14 });
  txt(s, "0.00 pp", { x: 8.0, y: 4.35, w: 4.4, h: 0.9, fontSize: 54, bold: true, color: C.accent1, fontFace: "Cambria" });
  txt(s, "gap in the Ricardian limit (infinite horizon, no constrained households), as theory requires", { x: 8.0, y: 5.3, w: 4.4, h: 0.8, fontSize: 14 });
  s.addNotes("The household budget, which the model does not impose, holds to 1e-15, a check that the accounting is consistent.");

  // 6 ------------------------------------------------------------ Kenya GIMF chart
  s = content("Results", `${D.name.toUpperCase()} EXAMPLE · GIMF-LITE`, `${D.name}'s net zero programme could lift GDP about ${Math.round(D.gimf["Debt financed"][25])}% by 2050`);
  const yrs = D.years.map(String);
  s.addChart(pres.charts.LINE, Object.keys(D.gimf).map((k) => ({ name: k, labels: yrs, values: D.gimf[k] })),
    Object.assign(chartBase("Real GDP, % deviation from no-programme path"), { x: 0.6, y: 1.7, w: 8.0, h: 4.9, valAxisLabelFormatCode: "0.0" }));
  const last = (k) => D.gimf[k][25];
  const share = Math.round(100 * (1 - last("No energy savings") / last("Debt financed")));
  const stats = [[`+${last("Debt financed").toFixed(1)}%`, "GDP in 2050, debt-financed public part"],
    [`${share}%`, "of that gain comes from energy savings"],
    [`${Math.round(D.gimf_debt["Debt financed"][25])}%`, "public debt / GDP in 2050 (50% today)"]];
  stats.forEach((st, i) => {
    card(s, 9.0, 1.7 + i * 1.65, 3.7, 1.45, `stat${i}`);
    txt(s, st[0], { x: 9.25, y: 1.8 + i * 1.65, w: 3.3, h: 0.7, fontSize: 36, bold: true, color: i === 1 ? C.accent2 : C.text1, fontFace: "Cambria" });
    txt(s, st[1], { x: 9.25, y: 2.5 + i * 1.65, w: 3.3, h: 0.55, fontSize: 13 });
  });
  txt(s, `Programme: NMIM v0.2 (${D.name}): public ${D.prog.pub[5]}%, private ${D.prog.prv[5]}% of GDP in 2030. Macro calibration illustrative.`,
    { x: 0.6, y: 6.65, w: 12, h: 0.3, fontSize: 10, color: C.accent4 });
  s.addNotes("The financing choice moves debt more than GDP: tax financing keeps debt falling with almost the same growth. Most of the gain comes from energy savings, so the savings assumption is the one to test first. Most of Kenya's public budget need is land-use and non-CO2 spending, which builds no capital, so the public part alone adds little to GDP.");

  // 7 ------------------------------------------------------------ DIGNAD
  s = content("Results", "DIGNAD-LITE · DISASTERS AND ADAPTATION", "Adapting before a disaster cuts the first-year loss by a sixth");
  const t20 = Array.from({ length: 20 }, (_, i) => String(i + 1));
  s.addChart(pres.charts.LINE, Object.keys(D.dignad).map((k) => ({ name: k, labels: t20, values: D.dignad[k] })),
    Object.assign(chartBase("Real GDP after a disaster, % deviation (years after the disaster)"), { x: 0.6, y: 1.7, w: 8.0, h: 4.9, valAxisLabelFormatCode: "0.0" }));
  card(s, 9.0, 1.7, 3.7, 4.9, "dignad_notes");
  txt(s, "What the scenarios show", { x: 9.25, y: 1.9, w: 3.3, h: 0.4, fontSize: 16, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, bullets([
    `Shock: 10% of capital destroyed, 3% TFP loss; GDP falls ${Math.abs(D.dignad["Reconstruction"][0]).toFixed(1)}%`,
    `Adaptation capital in place: ${D.dignad["Adaptation first + BBB"][0].toFixed(1)}% instead`,
    "Reconstruction speeds recovery but raises debt",
    "Building back better pays off only when the next disaster comes",
  ]), { x: 9.25, y: 2.4, w: 3.3, h: 4.0 });
  s.addNotes("DIGNAD-lite has a single surprise disaster, so the benefit of building back better shows only in the pre-disaster adaptation scenario. Recurrent disasters are the next extension.");

  // 8 ------------------------------------------------------------ DSF
  s = content("Results", "LIC-DSF-LITE", `${D.name}'s programme barely moves debt risk (illustrative)`);
  const th = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 }, fontSize: 13, align: "center" } });
  const tc = (t, o) => ({ text: String(t), options: Object.assign({ fontSize: 13, color: C.text2, fill: { color: C.background1 }, align: "center" }, o || {}) });
  const trows = [[th("Indicator (max over 10 years)"), th("Threshold"), th("Baseline"), th("With programme"), th("Worst stress"), th("Worst test")]];
  D.dsf.forEach((r) => trows.push([tc(r[0], { align: "left", bold: true }), tc(r[1]), tc(r[2]), tc(r[3], { bold: true }),
    tc(r[4], { color: r[4] > r[1] ? C.accent2 : C.text2, bold: r[4] > r[1] }), tc(r[5])]));
  s.addTable(trows, { x: 0.6, y: 1.75, w: 8.6, colW: [2.6, 1.1, 1.1, 1.35, 1.15, 1.3], rowH: 0.5, border: { type: "solid", pt: 0.75, color: HEX.line }, valign: "middle", margin: 0.06 });
  card(s, 9.6, 1.75, 3.1, 3.0, "rating");
  txt(s, "RISK RATING", { x: 9.85, y: 1.95, w: 2.6, h: 0.3, fontSize: 11, bold: true, color: C.accent2, charSpacing: 2 });
  txt(s, D.dsf_rating[1], { x: 9.85, y: 2.35, w: 2.6, h: 0.8, fontSize: 36, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, `unchanged from the baseline (${D.dsf_rating[0]}); breaches only under stress. Capacity ${D.ci[1]} (CI ${D.ci[0]}).`, { x: 9.85, y: 3.2, w: 2.6, h: 1.4, fontSize: 13 });
  txt(s, "Orange: stress value above the threshold. The baseline ratios are placeholders scaled to NMIM GDP, not Kenya's DSA; coefficients and thresholds as recalled from the 2018 Guidance Note, to be checked.",
    { x: 0.6, y: 5.4, w: 8.6, h: 0.8, fontSize: 11, italic: true, color: C.accent4 });
  s.addNotes("This is the mechanics, not a DSA: the rating depends on the baseline, which must come from the country's own DSA.");

  // 9 ------------------------------------------------------------ data
  pres.addSection({ title: "Next steps" });
  s = content("Next steps", "WHAT IT NEEDS", "A country calibration needs data most teams already hold");
  const cols = [["GIMF-lite", ["National accounts shares", "Revenue by tax, spending by type", "Public debt, net foreign assets", "Rates, inflation, monetary rule", "Constrained households (Findex)", "Return on public capital"]],
    ["DIGNAD-lite", ["Debt: concessional, domestic, commercial", "PIMA / PIE-X efficiency", "Grants and financing package", "Disaster damage (EM-DAT, PDNA)", "Adaptation cost and damage avoided", "Hand-to-mouth share"]],
    ["LIC-DSF-lite", ["10-20 year macro projections", "Loan-by-loan external debt service", "Terms of new borrowing", "CPIA, reserves, remittances", "Historical standard deviations", "Domestic debt"]]];
  cols.forEach((c, i) => {
    const x = 0.6 + i * 4.15;
    card(s, x, 1.75, 3.75, 3.7, `data${i}`);
    txt(s, c[0], { x: x + 0.3, y: 1.95, w: 3.2, h: 0.45, fontSize: 18, bold: true, color: i === 2 ? C.accent2 : C.accent1, fontFace: "Cambria" });
    txt(s, bullets(c[1]), { x: x + 0.3, y: 2.55, w: 3.2, h: 3.6 });
  });
  txt(s, "The investment programme itself already comes from NMIM for 186 countries.", { x: 0.6, y: 5.75, w: 12, h: 0.35, fontSize: 13, italic: true, color: C.accent4 });

  // 10 ------------------------------------------------------------ status
  s = content("Next steps", "STATUS", "v0.1 runs end to end; calibration is the next step");
  card(s, 0.6, 1.75, 5.9, 3.4, "done"); card(s, 6.8, 1.75, 5.9, 3.4, "notyet");
  s.addImage({ data: await icon(fa.FaCheckCircle, HEX.teal), x: 0.9, y: 2.0, w: 0.4, h: 0.4, altText: "Done" });
  txt(s, "Done", { x: 1.45, y: 2.0, w: 4.5, h: 0.4, fontSize: 18, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, bullets(["Three models, solved in seconds, on main in new-ai-tools", "Bridge to NMIM for any of its 186 countries", "Method note: equations and data needs", "Checks: Ricardian limit, Walras's law, horizon", "User guides for each model"]), { x: 0.9, y: 2.6, w: 5.3, h: 3.6 });
  s.addImage({ data: await icon(fa.FaExclamationTriangle, HEX.orange), x: 7.1, y: 2.0, w: 0.4, h: 0.4, altText: "Not yet" });
  txt(s, "Not yet", { x: 7.65, y: 2.0, w: 4.5, h: 0.4, fontSize: 18, bold: true, color: C.text1, fontFace: "Cambria" });
  txt(s, bullets(["Country calibrations (all illustrative today)", "DSF coefficients checked against the Guidance Note", "Energy savings treated as a value-added gain", "No carbon pricing, oil block or multi-country spillovers", "Recurrent disasters in DIGNAD-lite"]), { x: 7.1, y: 2.6, w: 5.3, h: 3.6 });

  // 11 ------------------------------------------------------------ next steps
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "Next steps" });
  s.addText("NEXT STEPS", { placeholder: "kicker" });
  s.addText("Pilot two or three countries, then build the front end", { placeholder: "title" });
  s.addText([
    { text: "1  Calibrate pilots with country teams' DSA and macro framework", options: { breakLine: true } },
    { text: "2  Check the DSF against the current Guidance Note and template", options: { breakLine: true } },
    { text: "3  Add carbon pricing revenue from CPAT", options: { breakLine: true } },
    { text: "4  Excel or dashboard view next to NMIM" },
  ], { placeholder: "body" });
  s.addNotes("Decision sought: which pilot countries, and who on the country teams provides the DSA and macro framework.");

  await pres.writeFile({ fileName: OUT });
  if (APPLY_THEME) { const { applyTheme } = require(APPLY_THEME); await applyTheme(OUT, THEME); }
  console.log("wrote", OUT);
})();
