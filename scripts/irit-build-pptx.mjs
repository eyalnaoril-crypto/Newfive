// אירית — בניית מצגת הנהלה לצי רכב (12 שנים, deep dive)
// פלט: output/<TS>-fleet-presentation-deep-dive-group-12yr.pptx

import PptxGenJS from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";

const NOW = new Date();
const ts = `${NOW.getFullYear()}${String(NOW.getMonth() + 1).padStart(2, "0")}${String(NOW.getDate()).padStart(2, "0")}-${String(NOW.getHours()).padStart(2, "0")}${String(NOW.getMinutes()).padStart(2, "0")}`;
const OUT = path.resolve("output", `${ts}-fleet-presentation-deep-dive-group-12yr.pptx`);

const C = {
  primary: "1F4E79",
  red: "C00000",
  green: "548235",
  gray: "7F7F7F",
  lightGray: "D9D9D9",
  bg: "FFFFFF",
  text: "262626",
  amber: "BF8F00",
};

const FONT = "Tahoma";
const fmt = (n) => "₪" + Math.round(n).toLocaleString("en-US");
const fmtNum = (n) => Math.round(n).toLocaleString("en-US");

// === Data (from gal deep-dive) ===
const YEARLY = {
  // year: [shmiraAcc, shmiraCost, migunAcc, migunCost]
  2014: [114, 216824, 336, 644089],
  2015: [100, 200411, 264, 490537],
  2016: [88, 370545, 149, 354906],
  2017: [114, 311069, 212, 413664],
  2018: [108, 182985, 285, 400404],
  2019: [135, 218704, 296, 464420],
  2020: [192, 336652, 264, 384915],
  2021: [177, 283326, 235, 298290],
  2022: [131, 270200, 231, 354083],
  2023: [146, 220284, 217, 225074],
  2024: [151, 275786, 286, 457795],
  2025: [154, 242317, 296, 455674],
  2026: [33, 31721, 45, 46137], // Q1 only
};
const TOT_SHMIRA_ACC = 1643;
const TOT_SHMIRA_COST = 3160824;
const TOT_MIGUN_ACC = 3116;
const TOT_MIGUN_COST = 4989989;
const TOT_ACC = TOT_SHMIRA_ACC + TOT_MIGUN_ACC; // 4759
const TOT_COST = TOT_SHMIRA_COST + TOT_MIGUN_COST; // 8,150,813
const YEARS_FULL = 12; // 2014-2025 full

// Top 10 branches across both divisions (combined)
const TOP_BRANCHES = [
  { name: "גוש דן (מיגון)", cost: 1447353 },
  { name: "פרוייקטים (מיגון)", cost: 529574 },
  { name: "אשדוד (שמירה)", cost: 500891 },
  { name: "חיפה (מיגון)", cost: 479492 },
  { name: 'פרוייקטים +acvs (מיגון)', cost: 459561 },
  { name: "רחובות (שמירה)", cost: 457703 },
  { name: "ירושלים (מיגון)", cost: 388570 },
  { name: "הוטלו (מיגון)", cost: 271033 },
  { name: 'פ"ת (שמירה)', cost: 259395 },
  { name: "חדרה (מיגון)", cost: 260989 },
];

const ACC_TYPES_SHMIRA = {
  "נהג החברה אשם": 419,
  "צד ג׳ / פריצה / חניה": 745,
  "נזק מרכב תחתון": 231,
  "נזקי סיום עיסקה": 251,
  "טוטל-לוס": 21,
};
const ACC_TYPES_MIGUN = {
  "נהג החברה אשם": 879,
  "צד ג׳ / פריצה / חניה": 1263,
  "נזק מרכב תחתון": 255,
  "נזקי סיום עיסקה": 661,
  "טוטל-לוס": 73,
};

const OUTLIERS = [
  { q: "2014-Q1", cost: 228026 },
  { q: "2014-Q2", cost: 181621 },
  { q: "2015-Q1", cost: 186188 },
  { q: "2019-Q1", cost: 182396 },
];

// ===== Build deck =====
const pres = new PptxGenJS();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.rtlMode = true;
pres.author = "אירית — סוכנת מצגות G1 Group";
pres.company = "G1 Group";
pres.title = "סקירת בטיחות צי 12 שנים";

const W = 13.333;
const H = 7.5;

// Footer helper
function addFooter(slide, cut) {
  slide.addShape("rect", { x: 0, y: H - 0.35, w: W, h: 0.35, fill: { color: C.primary }, line: { color: C.primary } });
  slide.addText(`רמת חתך: ${cut}   |   תקופה: 2014–Q1 2026   |   G1 Group — הופק ${String(NOW.getDate()).padStart(2, "0")}/${String(NOW.getMonth() + 1).padStart(2, "0")}/${NOW.getFullYear()}`, {
    x: 0.2, y: H - 0.33, w: W - 0.4, h: 0.3,
    fontFace: FONT, fontSize: 10, color: "FFFFFF",
    align: "right", rtlMode: true, valign: "middle",
  });
}

function titleBar(slide, title, sub) {
  // Top accent bar
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.08, fill: { color: C.primary }, line: { color: C.primary } });
  slide.addText(title, {
    x: 0.3, y: 0.18, w: W - 0.6, h: 0.7,
    fontFace: FONT, fontSize: 24, bold: true, color: C.primary,
    align: "right", rtlMode: true, valign: "middle",
  });
  if (sub) {
    slide.addText(sub, {
      x: 0.3, y: 0.85, w: W - 0.6, h: 0.4,
      fontFace: FONT, fontSize: 13, color: C.gray,
      align: "right", rtlMode: true, valign: "middle", italic: true,
    });
  }
}

// ===== 1. Cover =====
{
  const s = pres.addSlide();
  s.background = { color: C.primary };
  // Big logo block
  s.addShape("rect", { x: 0, y: 2.2, w: W, h: 0.05, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" } });
  s.addText("סקירת בטיחות צי", {
    x: 0.5, y: 2.5, w: W - 1, h: 1.0,
    fontFace: FONT, fontSize: 44, bold: true, color: "FFFFFF",
    align: "right", rtlMode: true,
  });
  s.addText("12 שנים של נתונים  |  שמירה מול מיגון  |  2014–Q1 2026", {
    x: 0.5, y: 3.6, w: W - 1, h: 0.6,
    fontFace: FONT, fontSize: 22, color: "FFFFFF",
    align: "right", rtlMode: true,
  });
  s.addText("G1 Group  |  מצגת הנהלה", {
    x: 0.5, y: 4.5, w: W - 1, h: 0.5,
    fontFace: FONT, fontSize: 18, color: "BFD4E6",
    align: "right", rtlMode: true,
  });
  s.addText(`הוכן ע״י אירית  |  ${String(NOW.getDate()).padStart(2, "0")}/${String(NOW.getMonth() + 1).padStart(2, "0")}/${NOW.getFullYear()}  |  רמת חתך: סך הקבוצה`, {
    x: 0.5, y: H - 0.9, w: W - 1, h: 0.4,
    fontFace: FONT, fontSize: 12, color: "BFD4E6",
    align: "right", rtlMode: true,
  });
}

// ===== 2. TOC =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "מה נראה היום", "5 תובנות, חריג נוכחי, ו-5 פעולות מדידות");

  const items = [
    "1.  תמונת מצב — צי הקבוצה ב-12 שנה",
    "2.  מיגון מייצר פי 1.9 תאונות ופי 1.58 בעלות",
    "3.  המגמה ארוכת-הטווח יורדת — אבל הפער בין החטיבות נשמר",
    "4.  גוש דן: ₪1.45M מצטבר — הסניף היקר במערכת",
    "5.  4 רבעונים חריגים — כולם במיגון, ב-2014–2019",
    "6.  פרופיל סוגי תאונה — שמירה=הכשרה, מיגון=חשיפה תפעולית",
    "7.  חריג Q1 2026 — ירושלים-מוקדים",
    "8.  המלצות פעולה — 5 צעדים מדידים",
  ];
  s.addText(items.map(t => ({ text: t, options: { fontFace: FONT, fontSize: 18, color: C.text, breakLine: true } })), {
    x: 0.8, y: 1.5, w: W - 1.6, h: 5.2,
    align: "right", rtlMode: true, valign: "top", paraSpaceAfter: 8,
  });
  addFooter(s, "סך הקבוצה");
}

// ===== 3. Snapshot — KPI cards =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "תמונת מצב — 12 שנה של צי הקבוצה", "סך הקבוצה (שמירה + מיגון) | 2014 – Q1 2026");

  // 4 KPI cards
  const cards = [
    { num: fmtNum(TOT_ACC), lbl: "סך תאונות 12 שנה", sub: `שמירה ${fmtNum(TOT_SHMIRA_ACC)}  |  מיגון ${fmtNum(TOT_MIGUN_ACC)}`, color: C.primary },
    { num: fmt(TOT_COST), lbl: "סך עלות תאונות 12 שנה", sub: `שמירה ${fmt(TOT_SHMIRA_COST)}  |  מיגון ${fmt(TOT_MIGUN_COST)}`, color: C.primary },
    { num: fmt(TOT_COST / YEARS_FULL), lbl: "ממוצע עלות שנתי", sub: `על בסיס 12 שנות נתונים מלאות`, color: C.primary },
    { num: "49", lbl: "רבעונים שדווחו", sub: `נתוני Q1 2026 חלקיים — בתהליך השלמה`, color: C.amber },
  ];
  const cardW = (W - 0.8 - 0.3 * 3) / 4;
  const yCard = 1.7;
  const hCard = 2.6;
  cards.forEach((c, i) => {
    const x = 0.4 + i * (cardW + 0.3);
    s.addShape("rect", { x, y: yCard, w: cardW, h: hCard, fill: { color: "F5F8FB" }, line: { color: c.color, width: 2 } });
    s.addShape("rect", { x, y: yCard, w: cardW, h: 0.15, fill: { color: c.color }, line: { color: c.color } });
    s.addText(c.num, { x: x + 0.1, y: yCard + 0.4, w: cardW - 0.2, h: 1.1, fontFace: FONT, fontSize: 32, bold: true, color: c.color, align: "center", rtlMode: true, valign: "middle" });
    s.addText(c.lbl, { x: x + 0.1, y: yCard + 1.5, w: cardW - 0.2, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: C.text, align: "center", rtlMode: true });
    s.addText(c.sub, { x: x + 0.1, y: yCard + 2.0, w: cardW - 0.2, h: 0.5, fontFace: FONT, fontSize: 11, color: C.gray, align: "center", rtlMode: true });
  });

  s.addText("מסר מרכזי: מערכת בטיחות יציבה — אבל מיגון מובילה ב-60% מהעלות ו-65% מהתאונות.", {
    x: 0.5, y: 4.7, w: W - 1, h: 0.6,
    fontFace: FONT, fontSize: 16, bold: true, color: C.primary, italic: true,
    align: "right", rtlMode: true,
  });

  addFooter(s, "סך הקבוצה");
}

// ===== 4. Insight 1: Migun dominance =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "מיגון מייצר פי 1.9 תאונות ופי 1.58 בעלות מהשמירה", "12 שנים מצטבר | סך הקבוצה");

  // Chart: horizontal bar comparison
  s.addChart(pres.ChartType.bar, [
    { name: "תאונות", labels: ["שמירה", "מיגון"], values: [TOT_SHMIRA_ACC, TOT_MIGUN_ACC] },
  ], {
    x: 0.5, y: 1.5, w: 6.0, h: 5.0,
    barDir: "bar",
    showTitle: true, title: "סך תאונות מצטבר", titleFontFace: FONT, titleFontSize: 14, titleColor: C.text,
    chartColors: [C.primary],
    showValue: true,
    dataLabelFontFace: FONT, dataLabelFontSize: 12, dataLabelColor: C.text,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10,
    showLegend: false,
  });

  s.addChart(pres.ChartType.bar, [
    { name: "עלות ₪", labels: ["שמירה", "מיגון"], values: [TOT_SHMIRA_COST, TOT_MIGUN_COST] },
  ], {
    x: 6.7, y: 1.5, w: 6.0, h: 5.0,
    barDir: "bar",
    showTitle: true, title: "סך עלות תאונות (₪)", titleFontFace: FONT, titleFontSize: 14, titleColor: C.text,
    chartColors: [C.red],
    showValue: true,
    dataLabelFontFace: FONT, dataLabelFontSize: 12, dataLabelColor: C.text,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10,
    showLegend: false,
  });

  s.addText("מיגון = נפח גבוה משמירה אך עלות לתאונה נמוכה יותר (₪1,601 לעומת ₪1,924). הסיכון = תדירות + חשיפה תפעולית.", {
    x: 0.5, y: 6.65, w: W - 1, h: 0.4,
    fontFace: FONT, fontSize: 13, color: C.gray, italic: true,
    align: "right", rtlMode: true,
  });

  addFooter(s, "סך הקבוצה — השוואה בין חטיבות");
}

// ===== 5. Insight 2: Trend (dual-line) =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "המגמה ארוכת-הטווח יורדת — אך הפער בין מיגון לשמירה נשמר", "עלות שנתית, ₪ | 2014–2025 (12 שנים מלאות)");

  const years = Object.keys(YEARLY).filter(y => y !== "2026").map(Number);
  const shmiraCost = years.map(y => YEARLY[y][1]);
  const migunCost = years.map(y => YEARLY[y][3]);

  s.addChart(pres.ChartType.line, [
    { name: "שמירה", labels: years.map(String), values: shmiraCost },
    { name: "מיגון", labels: years.map(String), values: migunCost },
  ], {
    x: 0.5, y: 1.5, w: W - 1, h: 4.8,
    chartColors: [C.green, C.red],
    lineSize: 3,
    lineDataSymbol: "circle", lineDataSymbolSize: 8,
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 13,
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 11,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10,
    showValue: false,
    valAxisLabelFormatCode: "#,##0",
  });

  s.addText("שיא היסטורי: מיגון 2014 (₪644K). שפל: מיגון 2023 (₪225K). 2024–2025 — עלייה מחודשת במיגון, דורש מעקב.", {
    x: 0.5, y: 6.45, w: W - 1, h: 0.4,
    fontFace: FONT, fontSize: 13, color: C.gray, italic: true,
    align: "right", rtlMode: true,
  });

  addFooter(s, "סך הקבוצה — מגמה רב-שנתית");
}

// ===== 6. Insight 3: Top branches =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "גוש דן: ₪1.45M ב-12 שנים — הסניף היקר במערכת", "Top 10 סניפים בעלות תאונות מצטברת | סך הקבוצה");

  const sorted = [...TOP_BRANCHES].sort((a, b) => a.cost - b.cost);
  s.addChart(pres.ChartType.bar, [
    { name: "עלות מצטברת ₪", labels: sorted.map(b => b.name), values: sorted.map(b => b.cost) },
  ], {
    x: 0.5, y: 1.4, w: W - 1, h: 5.3,
    barDir: "bar",
    chartColors: [C.primary],
    showValue: true,
    dataLabelFontFace: FONT, dataLabelFontSize: 10, dataLabelColor: C.text,
    dataLabelFormatCode: '"₪"#,##0',
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 11,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 9,
    valAxisLabelFormatCode: "#,##0",
    showLegend: false,
  });

  s.addText("גוש דן לבד = פי 2.7 מהסניף השני. ריכוז סיכון יוצא דופן — מצדיק בדיקת תוואי + בחינת מבנה ביטוח.", {
    x: 0.5, y: 6.8, w: W - 1, h: 0.3,
    fontFace: FONT, fontSize: 12, color: C.red, bold: true, italic: true,
    align: "right", rtlMode: true,
  });

  addFooter(s, "Top 10 סניפים | מצטבר 12 שנה");
}

// ===== 7. Insight 4: 4 outlier quarters =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "4 רבעונים חריגים (>2σ) — כולם במיגון, ב-2014–2019", "אין רבעון חריג מ-2020 ואילך — מערכת התייצבה");

  // Bar chart of outliers
  s.addChart(pres.ChartType.bar, [
    { name: "עלות מיגון ברבעון חריג (₪)", labels: OUTLIERS.map(o => o.q), values: OUTLIERS.map(o => o.cost) },
  ], {
    x: 0.5, y: 1.5, w: 7.5, h: 5.0,
    barDir: "col",
    chartColors: [C.red],
    showValue: true,
    dataLabelFontFace: FONT, dataLabelFontSize: 12, dataLabelColor: C.text,
    dataLabelFormatCode: '"₪"#,##0',
    catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12,
    valAxisLabelFontFace: FONT, valAxisLabelFontSize: 10,
    valAxisLabelFormatCode: "#,##0",
    showLegend: false,
  });

  // Right context box
  const bx = 8.4, by = 1.6, bw = 4.4, bh = 4.8;
  s.addShape("rect", { x: bx, y: by, w: bw, h: bh, fill: { color: "FDECEC" }, line: { color: C.red, width: 1 } });
  s.addText("מה משותף?", { x: bx + 0.2, y: by + 0.15, w: bw - 0.4, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: C.red, align: "right", rtlMode: true });
  const ctx = [
    "•  כל ה-4 ברבעון Q1 או Q2 (תחילת שנה)",
    "•  כולם בחטיבת מיגון בלבד",
    "•  שיא היסטורי: Q1 2014 — ₪228,026",
    "•  אין אירוע חריג מאז 2020 — סימן ליציבות מערכתית",
    "",
    "השלכה: מערכת הבקרה שהוטמעה ב-2019–2020 עובדת. צריך לשמר.",
  ];
  s.addText(ctx.map(t => ({ text: t, options: { fontFace: FONT, fontSize: 13, color: C.text, breakLine: true } })), {
    x: bx + 0.2, y: by + 0.7, w: bw - 0.4, h: bh - 0.8, align: "right", rtlMode: true, valign: "top", paraSpaceAfter: 6,
  });

  addFooter(s, "סך הקבוצה — חריגים סטטיסטיים");
}

// ===== 8. Insight 5: Accident-type profile (two donuts) =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "פרופיל סוגי תאונה: שמירה = הכשרה, מיגון = חשיפה תפעולית", "חלוקת תאונות לפי סוג | מצטבר 12 שנה");

  // Shmira donut
  s.addChart(pres.ChartType.doughnut, [
    { name: "שמירה", labels: Object.keys(ACC_TYPES_SHMIRA), values: Object.values(ACC_TYPES_SHMIRA) },
  ], {
    x: 0.3, y: 1.5, w: 6.3, h: 5.0,
    showTitle: true, title: "שמירה (אבטחה + ניקיון)", titleFontFace: FONT, titleFontSize: 14, titleColor: C.primary,
    chartColors: [C.red, C.primary, C.gray, C.amber, "B19CD9"],
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 11,
    showPercent: true, dataLabelFontFace: FONT, dataLabelFontSize: 10, dataLabelColor: "FFFFFF",
    holeSize: 50,
  });
  s.addChart(pres.ChartType.doughnut, [
    { name: "מיגון", labels: Object.keys(ACC_TYPES_MIGUN), values: Object.values(ACC_TYPES_MIGUN) },
  ], {
    x: 6.8, y: 1.5, w: 6.3, h: 5.0,
    showTitle: true, title: "מיגון (טכנולוגיות + מוקדים)", titleFontFace: FONT, titleFontSize: 14, titleColor: C.primary,
    chartColors: [C.red, C.primary, C.gray, C.amber, "B19CD9"],
    showLegend: true, legendPos: "b", legendFontFace: FONT, legendFontSize: 11,
    showPercent: true, dataLabelFontFace: FONT, dataLabelFontSize: 10, dataLabelColor: "FFFFFF",
    holeSize: 50,
  });

  s.addText("שמירה: 25% מהתאונות \"נהג אשם\" — סימן להכשרה. מיגון: 21% \"סיום עיסקה\" — חשיפת end-of-lease.", {
    x: 0.5, y: 6.65, w: W - 1, h: 0.4,
    fontFace: FONT, fontSize: 13, color: C.gray, italic: true,
    align: "right", rtlMode: true,
  });

  addFooter(s, "השוואת חטיבות — סוגי תאונה");
}

// ===== 9. Q1 2026 current outlier =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "Q1 2026 — תמונה ראשונית, נתונים חלקיים", "78 תאונות, ₪77,858 — נמוך משמעותית מממוצע רבעוני היסטורי");

  // Big number block
  const stats = [
    { lbl: "תאונות Q1 2026", val: "78", sub: "ממוצע רבעוני היסטורי: 99", color: C.green },
    { lbl: "עלות Q1 2026", val: fmt(77858), sub: `ממוצע רבעוני: ${fmt(166343)}`, color: C.green },
    { lbl: "ירידה מהממוצע", val: "-53%", sub: "אך נתון חלקי — דורש אימות", color: C.amber },
  ];
  const cardW = (W - 1 - 0.3 * 2) / 3;
  stats.forEach((c, i) => {
    const x = 0.5 + i * (cardW + 0.3);
    s.addShape("rect", { x, y: 1.7, w: cardW, h: 2.4, fill: { color: "F5F8FB" }, line: { color: c.color, width: 2 } });
    s.addText(c.val, { x, y: 1.85, w: cardW, h: 1.1, fontFace: FONT, fontSize: 40, bold: true, color: c.color, align: "center", rtlMode: true });
    s.addText(c.lbl, { x, y: 2.95, w: cardW, h: 0.5, fontFace: FONT, fontSize: 14, bold: true, color: C.text, align: "center", rtlMode: true });
    s.addText(c.sub, { x: x + 0.1, y: 3.45, w: cardW - 0.2, h: 0.5, fontFace: FONT, fontSize: 11, color: C.gray, align: "center", rtlMode: true });
  });

  // Warning box
  s.addShape("rect", { x: 0.5, y: 4.4, w: W - 1, h: 2.2, fill: { color: "FFF8E1" }, line: { color: C.amber, width: 1 } });
  s.addText("⚠  שאלות שעלינו לסגור לפני סיום Q1:", { x: 0.7, y: 4.55, w: W - 1.4, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: C.amber, align: "right", rtlMode: true });
  const qs = [
    "•  האם 0 תאונות בכמה סניפים = ירידה אמיתית או חוסר דיווח?",
    "•  עלות ק״מ נסיעה ב-Q1 2026 חסרה — חוסם חישוב KPI \"עלות לק״מ\"",
    "•  אם הירידה אמיתית, מה הסיבה? (תחילת השנה, תוכנית הכשרה, מזג אוויר?)",
  ];
  s.addText(qs.map(t => ({ text: t, options: { fontFace: FONT, fontSize: 13, color: C.text, breakLine: true } })), {
    x: 0.7, y: 5.0, w: W - 1.4, h: 1.5, align: "right", rtlMode: true, valign: "top", paraSpaceAfter: 6,
  });

  addFooter(s, "סך הקבוצה — Q1 2026 בלבד");
}

// ===== 10. Recommendations =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "5 פעולות מדידות לרבעון הקרוב", "אחראים מוצעים + אופק זמן ברור");

  const recs = [
    { num: "1", title: "בדיקת תוואי נסיעה ובחינת מבנה ביטוח — גוש דן", who: "מנהל חטיבת מיגון + ביטוח", when: "תוך 60 יום", color: C.red },
    { num: "2", title: "תוכנית הכשרת נהגים — חטיבת שמירה (25% \"נהג אשם\")", who: 'מנהל HR + יחידת הדרכה', when: "Q3 2026", color: C.amber },
    { num: "3", title: "הקצאת תקציב מניעת חשיפה למיגון — פריצות + סיום עיסקה", who: "סמנכ״ל תפעול + ביטוח", when: "תוך 90 יום", color: C.primary },
    { num: "4", title: "בירור Q1 2026 — אפס תאונות אמיתי או חוסר דיווח?", who: "אחראי דיווחי צי", when: "תוך 14 יום", color: C.amber },
    { num: "5", title: "השלמת נתוני ק״מ ל-2025–2026 — חוסם KPI \"עלות לק״מ\"", who: "אחראי מערכות תפעול", when: "תוך 30 יום", color: C.green },
  ];
  const y0 = 1.5;
  const rh = 1.05;
  recs.forEach((r, i) => {
    const y = y0 + i * rh;
    s.addShape("rect", { x: 0.5, y, w: W - 1, h: rh - 0.1, fill: { color: "F5F8FB" }, line: { color: r.color, width: 1 } });
    // Circle number
    s.addShape("ellipse", { x: W - 1.2, y: y + 0.18, w: 0.6, h: 0.6, fill: { color: r.color }, line: { color: r.color } });
    s.addText(r.num, { x: W - 1.2, y: y + 0.18, w: 0.6, h: 0.6, fontFace: FONT, fontSize: 22, bold: true, color: "FFFFFF", align: "center", valign: "middle" });
    // Title
    s.addText(r.title, { x: 0.8, y: y + 0.05, w: W - 2.2, h: 0.45, fontFace: FONT, fontSize: 15, bold: true, color: C.text, align: "right", rtlMode: true, valign: "middle" });
    // Meta
    s.addText(`אחראי: ${r.who}   |   אופק זמן: ${r.when}`, {
      x: 0.8, y: y + 0.5, w: W - 2.2, h: 0.4, fontFace: FONT, fontSize: 12, color: C.gray, align: "right", rtlMode: true, valign: "middle",
    });
  });

  addFooter(s, "המלצות פעולה — כל החטיבות");
}

// ===== 11. Call to action =====
{
  const s = pres.addSlide();
  s.background = { color: C.primary };
  s.addText("המגמה — חיובית.", { x: 0.5, y: 1.5, w: W - 1, h: 1.0, fontFace: FONT, fontSize: 44, bold: true, color: "FFFFFF", align: "right", rtlMode: true });
  s.addText("הסיכון — נשאר ממוקד: גוש דן + מיגון.", { x: 0.5, y: 2.7, w: W - 1, h: 1.0, fontFace: FONT, fontSize: 36, color: "FFFFFF", align: "right", rtlMode: true });
  s.addText("הצעד הבא — אישור 5 הפעולות.", { x: 0.5, y: 4.0, w: W - 1, h: 0.9, fontFace: FONT, fontSize: 30, bold: true, color: "FFD966", align: "right", rtlMode: true });

  s.addShape("rect", { x: 0.5, y: 5.5, w: W - 1, h: 0.05, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" } });
  s.addText("דיון נדרש: סדר עדיפויות לפעולות 1–5  |  אחראים סופיים  |  תקציב מניעה ל-Q3", {
    x: 0.5, y: 5.7, w: W - 1, h: 0.6, fontFace: FONT, fontSize: 16, color: "BFD4E6", align: "right", rtlMode: true, italic: true,
  });
  s.addText(`הוכן ע״י אירית — G1 Group  |  ${String(NOW.getDate()).padStart(2, "0")}/${String(NOW.getMonth() + 1).padStart(2, "0")}/${NOW.getFullYear()}`, {
    x: 0.5, y: H - 0.7, w: W - 1, h: 0.4, fontFace: FONT, fontSize: 11, color: "BFD4E6", align: "right", rtlMode: true,
  });
}

// ===== 12. Appendix — sources =====
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  titleBar(s, "נספח — מקורות נתונים והגדרות KPI", "שקיפות מלאה לקראת ביקורת");

  const lines = [
    "מקור ראשי: output/20260516-0850-fleet-shmira-deep-dive.xlsx (ניתוח גל, 13 לשוניות, 49 רבעונים).",
    "תקופת ניתוח: 2014-Q1 עד 2026-Q1 (12 שנים + רבעון חלקי).",
    "חטיבות: שמירה = אבטחה + ניקיון  |  מיגון = טכנולוגיות מיגון + חטיבת מוקדים.",
    "",
    "הגדרות KPI עיקריות:",
    "• עלות תאונות — סך עלות נזקים ברבעון, ללא עלות סיום עיסקה (טוטל-לוס נספר בנפרד).",
    "• עלות ממוצעת לתאונה = סך עלות / סך מס׳ תאונות באותה תקופה.",
    "• רבעון חריג = רבעון שעלותו > ממוצע + 2 סטיות תקן (כלל-מערכת).",
    "",
    "אזהרות איכות נתונים:",
    "• עמודות 2025–2026 — חלק מערכי ה-\"0\" עשויים להיות חוסר דיווח (סומנו לבירור).",
    "• ק״מ נסיעה Q1 2026 — חסר. KPI \"עלות לק״מ\" אינו זמין לתקופה זו.",
    "• 4 בלוקים היסטוריים — עלויות לא מפורטות. סומן N/A, לא הוחלף ב-0.",
  ];
  s.addText(lines.map(t => ({ text: t, options: { fontFace: FONT, fontSize: 13, color: t.startsWith("•") || t.startsWith("מקור") || t.startsWith("תקופת") || t.startsWith("חטיבות") ? C.text : C.text, bold: t === "הגדרות KPI עיקריות:" || t === "אזהרות איכות נתונים:", breakLine: true } })), {
    x: 0.6, y: 1.5, w: W - 1.2, h: 5.4, align: "right", rtlMode: true, valign: "top", paraSpaceAfter: 4,
  });

  addFooter(s, "נספח — סך הקבוצה");
}

// Save
await pres.writeFile({ fileName: OUT });
console.log("OK: " + OUT);
console.log("Slides: " + pres.slides.length);
