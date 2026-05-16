// scripts/bar-yoy-q1-report-v1_1.mjs
// בר — דוח Q1 YoY 2025→2026 — גרסה v1.1
// שינויים מ-v1.0:
//   1. כל פסקה שכללה רשימת סניפים פוצלה לשורות נפרדות עם מספור היררכי [פרק].[מספר].
//   2. כל פסקה שכללה רשימת תובנות/הערות פוצלה למספור.
//   3. תיקון הפנייה: "פרק 11" → "פרק 9" בתקציר מנהלים.
// פלט:  output/<YYYYMMDD-HHMM>-fleet-report-yoy-q1-2026-v1.1.docx

import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType,
  PageNumber, Header, Footer, TableOfContents,
  convertMillimetersToTwip, NumberFormat
} from 'docx';
import XLSX from 'xlsx';
import fs from 'fs';

// ---------- helpers ----------
const FONT = 'Tahoma';
const RTL  = true;

const wb = XLSX.readFile('output/20260516-0917-fleet-q1-yoy-2025-vs-2026.xlsx');
const sheet = (name) => XLSX.utils.sheet_to_json(wb.Sheets[name], {header:1, defval:null, blankrows:false});

const fmtInt = (n) => (typeof n === 'number') ? Math.round(n).toLocaleString('he-IL') : String(n ?? '');

function P(text, opts={}) {
  const { bold=false, size=22, align=AlignmentType.JUSTIFIED, heading, spacingAfter=120, spacingBefore=0, color } = opts;
  return new Paragraph({
    bidirectional: RTL,
    alignment: align,
    heading,
    spacing: { after: spacingAfter, before: spacingBefore, line: 276 },
    children: [new TextRun({ text, bold, size, font: FONT, rightToLeft: RTL, color })]
  });
}

function H(text, level) {
  const sizes = {1: 32, 2: 28, 3: 24};
  const heading = {1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3}[level];
  return new Paragraph({
    bidirectional: RTL,
    alignment: AlignmentType.RIGHT,
    heading,
    spacing: { before: level===1 ? 360 : 240, after: 120, line: 276 },
    children: [new TextRun({ text, bold: true, size: sizes[level], font: FONT, rightToLeft: RTL })]
  });
}

// תת-כותרת מודגשת בתוך פרק (לא heading-style — רק bold, יישור לימין, ריווח)
function SubH(text) {
  return new Paragraph({
    bidirectional: RTL,
    alignment: AlignmentType.RIGHT,
    spacing: { before: 160, after: 80, line: 276 },
    children: [new TextRun({ text, bold: true, size: 24, font: FONT, rightToLeft: RTL })]
  });
}

// שורת רשימה ממוספרת: "<chapter>.<idx>." בולד + טקסט רגיל
function NumItem(chapter, idx, text) {
  return new Paragraph({
    bidirectional: RTL,
    alignment: AlignmentType.RIGHT,
    indent: { start: 360 },
    spacing: { before: 40, after: 40, line: 276 },
    children: [
      new TextRun({ text: `${chapter}.${idx}. `, bold: true, size: 22, font: FONT, rightToLeft: RTL }),
      new TextRun({ text, size: 22, font: FONT, rightToLeft: RTL })
    ]
  });
}

// רשימה ממוספרת מתוך מערך פריטים — מחזיר Paragraphs[].
function NumberedList(chapter, startIdx, items) {
  return items.map((t, i) => NumItem(chapter, startIdx + i, t));
}

function cell(text, { header=false, width, align=AlignmentType.RIGHT } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    shading: header ? { type: ShadingType.CLEAR, fill: 'DCE6F1', color: 'auto' } : undefined,
    children: [new Paragraph({
      bidirectional: RTL,
      alignment: align,
      spacing: { before: 40, after: 40, line: 240 },
      children: [new TextRun({ text: String(text ?? ''), bold: header, size: 20, font: FONT, rightToLeft: RTL })]
    })]
  });
}

function tableFromRows(rows) {
  const trs = rows.map(r => new TableRow({
    tableHeader: r.header || false,
    children: r.cells.map(c => cell(c.text, { header: r.header, align: c.align })),
  }));
  return new Table({
    visuallyRightToLeft: true,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: trs,
    borders: {
      top:    { style: BorderStyle.SINGLE, size: 4, color: '808080' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: '808080' },
      left:   { style: BorderStyle.SINGLE, size: 4, color: '808080' },
      right:  { style: BorderStyle.SINGLE, size: 4, color: '808080' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'BFBFBF' },
      insideVertical:   { style: BorderStyle.SINGLE, size: 2, color: 'BFBFBF' },
    }
  });
}

// ---------- data ----------
const TOP5_SHM = [
  ['באר שבע', 7, 8704], ['מרכז', 2, 5200], ['עפולה', 5, 4559], ['עכו', 1, 4000], ['מפגשים', 4, 2320]
];
const TOP5_MIG = [
  ['ירושלים', 4, 13250], ['פרוייקטים +ACVS', 7, 9800], ['גוש דן', 15, 8142], ['ב"ש', 4, 4400], ['חיפה', 4, 3800]
];

const TYPE_MIX_SHM = [
  ['נהג החברה אשם', 5, 0.1515], ['צד ג\' / פריצה / חניה', 17, 0.5152],
  ['נזק מרכב תחתון', 8, 0.2424], ['נזקי סיום עיסקה', 3, 0.0909], ['טוטל-לוס', 0, 0]
];
const TYPE_MIX_MIG = [
  ['נהג החברה אשם', 11, 0.2444], ['צד ג\' / פריצה / חניה', 17, 0.3778],
  ['נזק מרכב תחתון', 4, 0.0889], ['נזקי סיום עיסקה', 9, 0.2000], ['טוטל-לוס', 4, 0.0889]
];
const TYPE_MIX_YOY = [
  ['נהג החברה אשם', 25, 0.2577, 16, 0.2051, -0.0526],
  ['צד ג\' / פריצה / חניה', 51, 0.5258, 34, 0.4359, -0.0899],
  ['נזק מרכב תחתון', 1, 0.0103, 12, 0.1538, 0.1435],
  ['נזקי סיום עיסקה', 19, 0.1959, 12, 0.1538, -0.0420],
  ['טוטל-לוס', 1, 0.0103, 4, 0.0513, 0.0410]
];

// ---------- build ----------
const today = new Date();
const datestr = `${String(today.getDate()).padStart(2,'0')}/${String(today.getMonth()+1).padStart(2,'0')}/${today.getFullYear()}`;
const children = [];

// Cover
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { before: 3000, after: 240 },
  children: [new TextRun({ text: 'דוח ניתוח רבעוני — בטיחות צי', bold: true, size: 48, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { after: 240 },
  children: [new TextRun({ text: 'Q1 2026 בהשוואה ל-Q1 2025', bold: true, size: 36, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { after: 1200 },
  children: [new TextRun({ text: 'קבוצת G1', size: 32, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: `תאריך הפקה: ${datestr}`, size: 24, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: 'גרסה: v1.1', size: 24, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ bidirectional: RTL, alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: 'מחבר: בר — סוכן דוחות צי', size: 24, font: FONT, rightToLeft: RTL })]}));
children.push(new Paragraph({ children: [], pageBreakBefore: true }));

// TOC
children.push(H('תוכן עניינים', 1));
children.push(new TableOfContents('תוכן עניינים', { hyperlink: true, headingStyleRange: '1-3' }));
children.push(new Paragraph({ children: [], pageBreakBefore: true }));

// 1. תקציר מנהלים
children.push(H('1. תקציר מנהלים', 1));
children.push(P(
  'ברבעון הראשון של 2026 רשמה קבוצת G1 ירידה רוחבית במספר התאונות ובעלותן ברמת הקבוצה. ' +
  'סך התאונות בצי עמד על 78 אירועים, לעומת 97 אירועים ברבעון הראשון של 2025, ירידה של 19.6%. ' +
  'סך עלות התאונות עמד על 77,858 ₪, לעומת 161,592 ₪ ברבעון המקביל, ירידה של 51.8%. ' +
  'הירידה בעלות חדה משמעותית מהירידה בנפח התאונות, ומצביעה על שיפור גם בחומרת האירועים הממוצעת.'
));

// 1.x — שלושה ממצאים מרכזיים (פוצל מפסקת "ראשית/שנית/שלישית")
children.push(SubH('שלושה ממצאים מרכזיים:'));
children.push(...NumberedList(1, 1, [
  'חטיבת שמירה (אבטחה וניקיון) הובילה את השיפור — ירידה של 32.7% במספר התאונות (49 → 33) וירידה של 62.8% בעלות (85,367 ₪ → 31,721 ₪).',
  'חטיבת מיגון (טכנולוגיות ומוקדים) שיפרה בעיקר את מימד העלות (-39.5%), בעוד שמספר התאונות ירד באופן מתון בלבד (-6.3%, 48 → 45).',
  'ברמת הסניפים נרשמו 15 סניפים שהשתפרו ביותר מ-20% מול 4 סניפים שהחמירו ביותר מ-20% — יחס של 3.75:1 לטובת השיפור.'
]));

// 1.x — שלושה צירי המלצה (פוצל מפסקת "(1)/(2)/(3)") + תיקון "פרק 11" → "פרק 9"
children.push(SubH('שלושה צירי המלצה עקרוניים:'));
children.push(...NumberedList(1, 4, [
  'שימור התשתית שהובילה לשיפור בחטיבת שמירה והפצתה לחטיבת מיגון — היעד המרכזי לשיפור ב-2026.',
  'טיפול ממוקד בארבעת הסניפים שהחמירו (באר שבע, פרוייקטים+ACVS, חיפה-מיגון, מרכז) ובסניף ירושלים-מיגון בו זוהתה עלות חריגה של 13,250 ₪ ברבעון.',
  'אימות איכות נתונים, ובפרט בדיקת חמשת הסניפים שאיפסו לחלוטין את הדיווח וקטגוריזציית "נזק מרכב תחתון" שזינקה מתאונה אחת ל-12.'
]));
children.push(P('פירוט מלא של ההמלצות מופיע בפרק 9.'));

// 2. רקע ומתודולוגיה
children.push(H('2. רקע ומתודולוגיה', 1));
children.push(P(
  'הדוח מסכם ניתוח השוואתי של אירועי בטיחות הצי בקבוצת G1 בין הרבעון הראשון של 2025 ל-Q1 2026. ' +
  'המקור הראשוני הוא קובץ "תאונות 1-2026" שעבר עיבוד ואיחוד בידי גל (Excel Fleet Analyst) ' +
  'לקובץ העבודה output/20260516-0917-fleet-q1-yoy-2025-vs-2026.xlsx, המכיל 11 לשוניות אגרגציה ' +
  'ברמות סניף, חטיבה וקבוצה.'
));
children.push(P(
  'הניתוח מתבסס על לשוניות הריכוז (שורות מסכמות פר סניף בכל רבעון), ולא על נתוני התאונות הגלמיים. ' +
  'מדדי המפתח שחושבו: מספר תאונות, סך עלות בש"ח, עלות ממוצעת לתאונה, ממוצע וחציון עלות לסניף, ' +
  'סטיית תקן וסף חריג סטטיסטי (mean+3σ), Δ מוחלט ו-Δ% מול Q1 2025, ותמהיל סוגי התאונות.'
));

// 2.x — מגבלות (פוצל מפסקת "(א)/(ב)/(ג)")
children.push(SubH('מגבלות עיקריות:'));
children.push(...NumberedList(2, 1, [
  'שדה ק"מ נסיעה ושדה מספר רכבים פעילים חסרים ב-Q1 2026 ולפיכך לא ניתן לחשב KPI עלות-לק"מ או עלות-לרכב.',
  'חמישה סניפים שדיווחו ב-2025 מציגים אפס תאונות ואפס עלות ב-Q1 2026 — האם השיפור אמיתי או חוסר דיווח דורש אימות.',
  'הקטגוריה "נזק מרכב תחתון" עלתה מ-1 ל-12 אירועים — דורש בירור האם מדובר בקטגוריזציה חדשה.'
]));

// 3. תמונת מצב — קבוצה
children.push(H('3. תמונת מצב — סך הקבוצה', 1));
children.push(P(
  'ברמת הקבוצה ב-Q1 2026 נרשמו 78 תאונות בעלות מצטברת של 77,858 ₪. עלות ממוצעת לתאונה ' +
  'עמדה על 998 ₪. הצי נחלק לשתי חטיבות: שמירה עם 21 סניפים פעילים ו-33 תאונות, ומיגון ' +
  'עם 20 סניפים פעילים ו-45 תאונות.'
));
children.push(tableFromRows([
  { header: true, cells: [
    {text:'מדד'},{text:'Q1 2025'},{text:'Q1 2026'},{text:'Δ מוחלט'},{text:'Δ%'}
  ]},
  { cells: [{text:'תאונות — שמירה'},{text:'49'},{text:'33'},{text:'−16'},{text:'−32.7%'}]},
  { cells: [{text:'תאונות — מיגון'},{text:'48'},{text:'45'},{text:'−3'},{text:'−6.3%'}]},
  { cells: [{text:'תאונות — קבוצה'},{text:'97'},{text:'78'},{text:'−19'},{text:'−19.6%'}]},
  { cells: [{text:'עלות — שמירה (₪)'},{text:'85,367'},{text:'31,721'},{text:'−53,646'},{text:'−62.8%'}]},
  { cells: [{text:'עלות — מיגון (₪)'},{text:'76,225'},{text:'46,137'},{text:'−30,088'},{text:'−39.5%'}]},
  { cells: [{text:'עלות — קבוצה (₪)'},{text:'161,592'},{text:'77,858'},{text:'−83,734'},{text:'−51.8%'}]},
]));
children.push(P(''));
children.push(P(
  'הקריאה ברמת הקבוצה היא של שיפור רוחבי. עם זאת, הירידה בעלות (-51.8%) חדה משמעותית ' +
  'מהירידה בנפח (-19.6%), כלומר התאונות ב-Q1 2026 קלות יותר בממוצע. עלות ממוצעת לתאונה ' +
  'ירדה מ-1,666 ₪ ל-998 ₪ — ירידה של כ-40%.'
));

// 4. חטיבת שמירה
children.push(H('4. ניתוח חטיבת שמירה (אבטחה וניקיון)', 1));
children.push(H('4.1 KPIs ראשיים', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'מדד'},{text:'ערך Q1 2026'}]},
  { cells: [{text:'סה"כ תאונות'},{text:'33'}]},
  { cells: [{text:'סה"כ עלות (₪)'},{text:'31,721'}]},
  { cells: [{text:'סניפים פעילים'},{text:'21'}]},
  { cells: [{text:'עלות ממוצעת לתאונה (₪)'},{text:'961'}]},
  { cells: [{text:'ממוצע עלות לסניף (₪)'},{text:'2,643'}]},
  { cells: [{text:'חציון עלות לסניף (₪)'},{text:'1,698'}]},
]));
children.push(P(''));
children.push(P(
  'ברמת חטיבת שמירה נרשם השיפור הגדול בקבוצה: ירידה של 32.7% במספר התאונות וירידה של ' +
  '62.8% בעלות. הפער בין ממוצע (2,643 ₪) לחציון (1,698 ₪) לסניף מעיד על התפלגות מוטה — ' +
  'מספר קטן של סניפים מרכז חלק ניכר מהעלות.'
));

children.push(H('4.2 חמשת הסניפים היקרים — Q1 2026', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'סניף'},{text:'תאונות'},{text:'עלות (₪)'}]},
  ...TOP5_SHM.map(([b,a,c]) => ({ cells: [{text:b},{text:fmtInt(a)},{text:fmtInt(c)}] }))
]));
children.push(P(''));

children.push(H('4.3 תמהיל סוגי תאונה', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'סוג'},{text:'כמות'},{text:'שיעור'}]},
  ...TYPE_MIX_SHM.map(([t,q,p]) => ({ cells: [{text:t},{text:fmtInt(q)},{text:(p*100).toFixed(1)+'%'}] }))
]));
children.push(P(''));
children.push(P(
  'הקטגוריה השכיחה היא "צד ג\' / פריצה / חניה" (51.5%, 17 אירועים) — אירועים שאינם תלויים ' +
  'בהכרח בנהג. הקטגוריה "נהג החברה אשם" עומדת על 15.2% (5 אירועים), ירידה משיעור של 25.8% ' +
  'בכלל הקבוצה ב-Q1 2025.'
));

// 5. חטיבת מיגון
children.push(H('5. ניתוח חטיבת מיגון (טכנולוגיות ומוקדים)', 1));
children.push(H('5.1 KPIs ראשיים', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'מדד'},{text:'ערך Q1 2026'}]},
  { cells: [{text:'סה"כ תאונות'},{text:'45'}]},
  { cells: [{text:'סה"כ עלות (₪)'},{text:'46,137'}]},
  { cells: [{text:'סניפים פעילים'},{text:'20'}]},
  { cells: [{text:'עלות ממוצעת לתאונה (₪)'},{text:'1,025'}]},
  { cells: [{text:'ממוצע עלות לסניף (₪)'},{text:'4,614'}]},
  { cells: [{text:'חציון עלות לסניף (₪)'},{text:'3,300'}]},
]));
children.push(P(''));
children.push(P(
  'בחטיבת מיגון נרשמה ירידה מתונה במספר התאונות (-6.3%) אך ירידה משמעותית בעלות (-39.5%). ' +
  'עלות ממוצעת לתאונה (1,025 ₪) גבוהה ב-6.7% מזו של חטיבת שמירה (961 ₪). ' +
  'ממוצע עלות לסניף בחטיבה (4,614 ₪) גבוה כמעט פי שניים מחטיבת שמירה (2,643 ₪).'
));

children.push(H('5.2 חמשת הסניפים היקרים — Q1 2026', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'סניף'},{text:'תאונות'},{text:'עלות (₪)'}]},
  ...TOP5_MIG.map(([b,a,c]) => ({ cells: [{text:b},{text:fmtInt(a)},{text:fmtInt(c)}] }))
]));
children.push(P(''));
children.push(P(
  'ברמת סניף בולט ירושלים (מיגון) — 4 תאונות בעלות 13,250 ₪, כלומר 3,313 ₪ לתאונה בממוצע, ' +
  'פי 3.2 מהממוצע הקבוצתי. עלות זו דורשת בירור.'
));

children.push(H('5.3 תמהיל סוגי תאונה', 2));
children.push(tableFromRows([
  { header: true, cells: [{text:'סוג'},{text:'כמות'},{text:'שיעור'}]},
  ...TYPE_MIX_MIG.map(([t,q,p]) => ({ cells: [{text:t},{text:fmtInt(q)},{text:(p*100).toFixed(1)+'%'}] }))
]));
children.push(P(''));
children.push(P(
  'בחטיבת מיגון שיעור גבוה יותר של "נהג החברה אשם" (24.4% מול 15.2% בשמירה) ושל "טוטל-לוס" ' +
  '(8.9% מול 0% בשמירה). דפוס זה מצביע על אירועים חמורים יותר בממוצע ועל מקום לחיזוק הכשרה ' +
  'של עובדי הצי בחטיבה.'
));

// 6. השוואה בין החטיבות
children.push(H('6. השוואה בין החטיבות', 1));
children.push(tableFromRows([
  { header: true, cells: [{text:'מדד'},{text:'שמירה'},{text:'מיגון'},{text:'הפרש'}]},
  { cells: [{text:'תאונות'},{text:'33'},{text:'45'},{text:'−12'}]},
  { cells: [{text:'עלות (₪)'},{text:'31,721'},{text:'46,137'},{text:'−14,416'}]},
  { cells: [{text:'סניפים פעילים'},{text:'21'},{text:'20'},{text:'+1'}]},
  { cells: [{text:'עלות ממוצעת לתאונה (₪)'},{text:'961'},{text:'1,025'},{text:'−64'}]},
  { cells: [{text:'Δ% תאונות YoY'},{text:'−32.7%'},{text:'−6.3%'},{text:'—'}]},
  { cells: [{text:'Δ% עלות YoY'},{text:'−62.8%'},{text:'−39.5%'},{text:'—'}]},
]));
children.push(P(''));

// 6.x — שלוש תובנות השוואתיות (פוצל מפסקה אחת)
children.push(SubH('שלוש תובנות השוואתיות:'));
children.push(...NumberedList(6, 1, [
  'חטיבת שמירה היא המנוע של השיפור הקבוצתי ברבעון הנוכחי — מובילה הן בקצב הירידה בתאונות והן בקצב הירידה בעלות.',
  'חטיבת מיגון משתפרת אף היא, אך באופן מתון יותר, ובעיקר בממד העלות (-39.5%) ופחות בנפח (-6.3%).',
  'מבחינת הרכב התאונות, חטיבת מיגון רגישה יותר לאירועים בהם הנהג אשם או שמסתיימים בטוטל-לוס.'
]));

children.push(H('6.1 תמהיל סוגי תאונה — קבוצה YoY', 2));
children.push(tableFromRows([
  { header: true, cells: [
    {text:'סוג'},{text:'Q1 2025'},{text:'2025 %'},{text:'Q1 2026'},{text:'2026 %'},{text:'Δ%'}
  ]},
  ...TYPE_MIX_YOY.map(([t,q1,p1,q2,p2,d]) => ({ cells: [
    {text:t},{text:fmtInt(q1)},{text:(p1*100).toFixed(1)+'%'},
    {text:fmtInt(q2)},{text:(p2*100).toFixed(1)+'%'},
    {text:(d>=0?'+':'−')+(Math.abs(d)*100).toFixed(1)+' נק"א'}
  ]}))
]));
children.push(P(''));
children.push(P(
  'הירידות הבולטות: "צד ג\' / פריצה / חניה" (-9 נקודות אחוז) ו"נהג החברה אשם" (-5.3 נקודות). ' +
  'לעומת זאת, הקטגוריה "נזק מרכב תחתון" קפצה מ-1 ל-12 אירועים — דורש בדיקה האם מדובר בשינוי ' +
  'תופעתי או בקטגוריזציה חדשה שלא הייתה קיימת ב-2025.'
));

// 7. ממצאים וחריגים
children.push(H('7. ממצאים וחריגים', 1));

children.push(H('7.1 סניפים שהחמירו ב-20% או יותר', 2));
children.push(tableFromRows([
  { header: true, cells: [
    {text:'סניף'},{text:'חטיבה'},{text:'תאונות 25'},{text:'תאונות 26'},{text:'עלות 25 (₪)'},{text:'עלות 26 (₪)'}
  ]},
  { cells: [{text:'מרכז'},{text:'שמירה'},{text:'1'},{text:'2'},{text:'נתון חסר'},{text:'5,200'}]},
  { cells: [{text:'באר שבע'},{text:'שמירה'},{text:'5'},{text:'7'},{text:'5,709'},{text:'8,704'}]},
  { cells: [{text:'פרוייקטים +ACVS'},{text:'מיגון'},{text:'5'},{text:'7'},{text:'12,660'},{text:'9,800'}]},
  { cells: [{text:'חיפה'},{text:'מיגון'},{text:'3'},{text:'4'},{text:'4,690'},{text:'3,800'}]},
]));
children.push(P(''));

// 7.1.x — פירוט הסניפים שהחמירו (פוצל לרשימה ממוספרת)
children.push(SubH('פירוט הסניפים שהחמירו:'));
children.push(...NumberedList(7, 1, [
  'באר שבע (שמירה) — 5 → 7 תאונות; עלות 5,709 ₪ → 8,704 ₪ (+52.5%).',
  'פרוייקטים+ACVS (מיגון) — 5 → 7 תאונות; עלות 12,660 ₪ → 9,800 ₪ (-22.6%).',
  'חיפה (מיגון) — 3 → 4 תאונות; עלות 4,690 ₪ → 3,800 ₪ (-19.0%).',
  'מרכז (שמירה) — 1 → 2 תאונות; עלות 5,200 ₪ ב-Q1 2026 (נתון 2025 חסר).',
  'ירושלים (מיגון) — אף שמספר התאונות ירד מ-9 ל-4, העלות עלתה ב-48% ל-13,250 ₪ — חריג שדורש בירור.'
]));

children.push(H('7.2 סניפים שאיפסו דיווח — דורש אימות', 2));
children.push(P('שישה סניפים דיווחו על תאונות ב-Q1 2025 ועל אפס תאונות ב-Q1 2026:'));

children.push(SubH('סניפים שעברו לאפס תאונות ב-Q1 2026 (חטיבת שמירה):'));
children.push(...NumberedList(7, 6, [
  'כפר סבא — 6 → 0 תאונות; עלות 17,891 ₪ → 0 ₪.',
  'פ"ת — 8 → 0 תאונות; עלות 13,649 ₪ → 0 ₪.',
  'ראש פינה — 3 → 0 תאונות; עלות 3,665 ₪ → 0 ₪.',
  'ראש פינה ניקיון — 1 → 0 תאונות; עלות 1,000 ₪ → 0 ₪.',
  'מטה — 2 → 0 תאונות.',
  'אבטחה מדלגת — 1 → 0 תאונות; עלות 14,493 ₪ → 0 ₪.'
]));

children.push(SubH('שתי קריאות אפשריות לתופעה:'));
children.push(...NumberedList(7, 12, [
  'שיפור אמיתי בעקבות פעולה ייעודית (הכשרה, החלפת נהגים, שינוי תוואי).',
  'חוסר דיווח או שינוי שיוך ארגוני — נדרש אימות עם אחראי הסניף לפני שהמספרים יישלחו הלאה כתוצאה תפעולית.'
]));

children.push(H('7.3 דגלי איכות נתונים', 2));
children.push(SubH('שני דגלי איכות נתונים מרכזיים:'));
children.push(...NumberedList(7, 14, [
  'קטגוריית "נזק מרכב תחתון" עברה מאירוע יחיד ב-Q1 2025 ל-12 אירועים ב-Q1 2026 (8 בשמירה, 4 במיגון). היקף הקפיצה (פי 12) חשוד יותר מתופעה אמיתית ופחות מקטגוריזציה חדשה או שינוי בהגדרת השדה. דרושה התייעצות עם מנהל הצי.',
  'שדה הקילומטראז\' הרבעוני ושדה מספר הרכבים הפעילים מופיעים כאפס לכל הסניפים ב-Q1 2026, ולכן KPI עלות-לק"מ ועלות-לרכב לא חושב. השלמת נתון זה תאפשר ניתוח עומק נוסף ברבעון הבא.'
]));

// 8. מסקנות
children.push(H('8. מסקנות', 1));
children.push(SubH('ארבע מסקנות מרכזיות:'));
children.push(...NumberedList(8, 1, [
  'הצי מציג שיפור מובהק ורחב היקף ברבעון הראשון של 2026 מול הרבעון המקביל אשתקד. הירידה של 19.6% בנפח התאונות מלווה בירידה חדה יותר של 51.8% בעלות — הן פחות תאונות והן פחות חמורות. השיפור לא נשען על סניף בודד אלא נפרש על 15 סניפים שונים.',
  'חטיבת שמירה היא סיפור ההצלחה של הרבעון. הירידה של 32.7% בתאונות והירידה של 62.8% בעלות הן השיפור הגדול ברמת חטיבה. הצמצום בקטגוריית "נהג החברה אשם" (15.2% בלבד מתאונות החטיבה) מעיד שגורמים ניהוליים — הכשרה, אכיפת מהירויות, או שינוי תוואי נסיעה — תורמים לתוצאה. ההמלצה היא לתעד את הפעולות שננקטו ולהפיצן לחטיבות הנוספות.',
  'חטיבת מיגון היא היעד הבא לסגירת הפער. אף ששיפור העלות בה משמעותי (-39.5%), נפח התאונות ירד רק ב-6.3% והרכב התאונות חמור יותר (24.4% "נהג אשם" ו-8.9% "טוטל-לוס"). מיקוד בהכשרת נהגים ובניהול סיכון בחטיבה זו צפוי להניב את התשואה הגבוהה ביותר ב-2026.',
  'קיימים דגלים אדומים שאסור להתעלם מהם: סניף ירושלים-מיגון עם עלות ממוצעת חריגה של 3,313 ₪ לתאונה; ארבעה סניפים שהחמירו ביותר מ-20%; חמישה סניפים שאיפסו את הדיווח שלהם; וקטגוריזציה חשודה של "נזק מרכב תחתון". אלה דורשים מענה תפעולי לפני שהדוח הופך לבסיס להחלטות תקציביות.'
]));

// 9. המלצות
children.push(H('9. המלצות פעולה', 1));
const recs = [
  ['תיעוד תוכנית ההצלחה של חטיבת שמירה', 'תיעוד מובנה של הפעולות שננקטו בחטיבת שמירה במהלך Q4 2025 / Q1 2026 (הכשרות, נהלים, אכיפה) שהובילו לירידה של 32.7%. שיתוף עם הנהלת חטיבת מיגון כתשתית להעתקה.'],
  ['תוכנית הכשרה ייעודית בחטיבת מיגון', 'מיקוד בקטגוריית "נהג החברה אשם" (24.4% בחטיבה). יעד מדיד: הורדה ל-18% עד סוף Q3 2026.'],
  ['בירור עלות חריגה — ירושלים-מיגון', 'אימות מקור העלות של 13,250 ₪ ב-4 תאונות בלבד. הכוונה אפשרית: אירוע חמור יחיד, או חיוב מצטבר שנגרר. החלטה האם מדובר בחריג חד-פעמי או דפוס.'],
  ['טיפול ממוקד בסניפים שהחמירו', 'באר שבע (שמירה), פרוייקטים+ACVS (מיגון), חיפה (מיגון), מרכז (שמירה) — שיחת חקר מובנית עם מנהל הסניף, מיפוי גורמים, יעד שיפור ל-Q2 2026.'],
  ['אימות הסניפים שאיפסו', 'התקשרות עם 5 סניפי שמירה (כפר סבא, פ"ת, ראש פינה, ראש פינה ניקיון, מטה) ו-"אבטחה מדלגת" — לאמת האם האפס נכון, או נובע מחוסר דיווח / שינוי שיוך.'],
  ['בירור קטגוריזציית "נזק מרכב תחתון"', 'הקפיצה מ-1 ל-12 אירועים חשודה. מומלץ לבדוק עם מנהל הצי האם נוסף קוד תאונה חדש, או שמדובר בשינוי דפוס נסיעה אמיתי.'],
  ['השלמת נתוני ק"מ ומס\' רכבים ל-Q1 2026', 'ללא נתונים אלה לא ניתן לחשב עלות-לק"מ ועלות-לרכב — שניים מ-KPI הליבה. יעד: השלמה לפני סגירת ניתוח Q2.'],
  ['הצבת חטיבת מיגון כיעד שיפור מרכזי ל-2026', 'הגדרת KPI שנתי: ירידה של 20% בתאונות חטיבת מיגון (45 → 36 ברבעון), ירידה של 30% בעלות. מעקב רבעוני.'],
  ['בדיקה היסטורית: האם השיפור YoY רצוף או חד-פעמי', 'הרחבת ההשוואה לרבעונים 2024-Q1, 2023-Q1, וכך הלאה — לזהות אם זו מגמה רב-שנתית או "אפקט בסיס".'],
];
recs.forEach(([title, body], i) => {
  children.push(new Paragraph({
    bidirectional: RTL,
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 120, after: 80, line: 276 },
    children: [new TextRun({ text: `9.${i+1}. ${title}. `, bold: true, size: 22, font: FONT, rightToLeft: RTL }),
               new TextRun({ text: body, size: 22, font: FONT, rightToLeft: RTL })]
  }));
});

// 10. נספח
children.push(new Paragraph({ children: [], pageBreakBefore: true }));
children.push(H('10. נספח', 1));
children.push(H('10.1 מילון KPI', 2));
const kpiDict = sheet('מילון KPI').slice(1).filter(r => r[0]);
children.push(tableFromRows([
  { header: true, cells: [{text:'מדד'},{text:'הגדרה / נוסחה'},{text:'יחידה'}]},
  ...kpiDict.map(r => ({ cells: [{text:String(r[0]??'')},{text:String(r[1]??'')},{text:String(r[2]??'')}] }))
]));
children.push(P(''));

children.push(H('10.2 מקורות נתונים', 2));
const sources = sheet('מקורות נתונים').slice(1).filter(r => r[0]);
children.push(tableFromRows([
  { header: true, cells: [{text:'פריט'},{text:'פירוט'}]},
  ...sources.map(r => ({ cells: [{text:String(r[0]??'')},{text:String(r[1]??'')}] }))
]));

// ---------- Document ----------
const doc = new Document({
  creator: 'בר — סוכן דוחות צי, G1-Group',
  title: 'דוח רבעוני בטיחות צי Q1 2026 — v1.1',
  styles: {
    default: {
      document: { run: { font: FONT, size: 22 }, paragraph: { spacing: { line: 276 } } },
      heading1: { run: { font: FONT, size: 32, bold: true } },
      heading2: { run: { font: FONT, size: 28, bold: true } },
      heading3: { run: { font: FONT, size: 24, bold: true } },
    },
  },
  numbering: { config: [] },
  features: { updateFields: true },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertMillimetersToTwip(25),
          bottom: convertMillimetersToTwip(25),
          left: convertMillimetersToTwip(25),
          right: convertMillimetersToTwip(25),
        },
        pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
      },
      bidi: true,
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          bidirectional: RTL,
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'דוח רבעוני בטיחות צי | Q1 2026 vs Q1 2025    G1-Group    v1.1', size: 18, font: FONT, rightToLeft: RTL, color: '707070' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          bidirectional: RTL,
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: ['עמוד ', PageNumber.CURRENT, ' מתוך ', PageNumber.TOTAL_PAGES], size: 18, font: FONT, rightToLeft: RTL })]
        })]
      })
    },
    children
  }]
});

const pad = n => String(n).padStart(2,'0');
const ts = `${today.getFullYear()}${pad(today.getMonth()+1)}${pad(today.getDate())}-${pad(today.getHours())}${pad(today.getMinutes())}`;
const outPath = `output/${ts}-fleet-report-yoy-q1-2026-v1.1.docx`;

const buf = await Packer.toBuffer(doc);
fs.writeFileSync(outPath, buf);
console.log('Wrote:', outPath, '(' + (buf.length/1024).toFixed(1) + ' KB)');
