// scripts/barak-pdf-yoy-q1.mjs
// ברק — הפקת PDF פנימי v1.1 משילוב של:
//   1) דוח Word v1.1 (mammoth -> טקסט מובנה)
//   2) נספח עם 4-5 גרפים מנתוני YoY Q1 2025 vs Q1 2026
// פלט: output/<YYYYMMDD-HHMM>-fleet-report-yoy-q1-2026-v1.1.pdf
// תכונות: Header/Footer/Watermark, פונט Tahoma מוטבע, RTL לטקסט עברית.

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import PDFDocument from 'pdfkit';
import mammoth from 'mammoth';
import ExcelJS from 'exceljs';
import { ChartJSNodeCanvas } from 'chartjs-node-canvas';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const DOCX = path.join(ROOT, 'output', '20260516-0951-fleet-report-yoy-q1-2026-v1.1.docx');
const XLSX_SRC = path.join(ROOT, 'output', '20260516-0917-fleet-q1-yoy-2025-vs-2026.xlsx');

const TAHOMA_REG = 'C:/Windows/Fonts/tahoma.ttf';
const TAHOMA_BD  = 'C:/Windows/Fonts/tahomabd.ttf';

const stamp = (() => {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return {
    file: `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`,
    iso:  `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`
  };
})();

const OUT = path.join(ROOT, 'output', `${stamp.file}-fleet-report-yoy-q1-2026-v1.1.pdf`);

// ---------- RTL helper ----------
// pdfkit לא יודע BIDI. נשתמש ב-`features` (OpenType) כדי להחיל RTL לטקסט עברית.
// pdfkit תומך ב-options.features=['rtla'] ו-options.direction. אבל הדרך הבטוחה ביותר
// היא שימוש ב-options.features=['rtla'] + יישור לימין. עברית בלי ניקוד מתחברת אוטומטית
// כי מדובר באותיות נפרדות; הסדר הויזואלי הוא היפוך הסדר הלוגי.
// פתרון: היפוך ידני של רצפי תווים עברית בכל שורה (מילים בעברית בלבד או מעורב).
function reverseLine(text) {
  if (!text) return text;
  // פירוק לפיסות לפי רווחים; היפוך רצפי מילים עבריות עם שמירת LTR runs.
  // אך פשוט יותר: היפוך ברמת מילים אם השורה רובה עברית.
  const hasHeb = /[֐-׿]/.test(text);
  if (!hasHeb) return text;
  // אסטרטגיה: פיצול ל-tokens, היפוך סדר ה-tokens, ובתוך כל token עברי — היפוך תווים.
  // אם הtoken מכיל אותיות לטיניות/מספרים — שמור כסדרו.
  // לבסוף — חזור על השורה הפוכה.
  const tokens = text.split(/(\s+)/); // משמר רווחים
  // היפוך תווים בתוך כל token עברי
  const fixed = tokens.map(tok => {
    if (/^\s+$/.test(tok)) return tok;
    if (/[֐-׿]/.test(tok) && !/[A-Za-z]/.test(tok)) {
      // עברית בלבד (אולי עם פיסוק/מספרים) — היפוך
      // מספרים בעברית צריכים להישאר LTR; נטפל בקבוצות.
      return reverseHebrewSegment(tok);
    }
    return tok;
  });
  // היפוך סדר ה-tokens (כך שמילה אחרונה תופיע ראשונה ויוצג נכון מימין)
  return fixed.reverse().join('');
}

function reverseHebrewSegment(s) {
  // היפוך תווים עם שמירת מספרים/אותיות לטיניות רצופות (LTR runs) בסדרם.
  // נפרק ל-runs: hebrew chars, digits, latin chars, punctuation
  const runs = [];
  let i = 0;
  while (i < s.length) {
    const ch = s[i];
    if (/[֐-׿]/.test(ch)) {
      let j = i;
      while (j < s.length && /[֐-׿]/.test(s[j])) j++;
      runs.push({ type: 'heb', text: s.slice(i, j) });
      i = j;
    } else if (/[0-9]/.test(ch)) {
      let j = i;
      while (j < s.length && /[0-9.,]/.test(s[j])) j++;
      runs.push({ type: 'num', text: s.slice(i, j) });
      i = j;
    } else if (/[A-Za-z]/.test(ch)) {
      let j = i;
      while (j < s.length && /[A-Za-z]/.test(s[j])) j++;
      runs.push({ type: 'lat', text: s.slice(i, j) });
      i = j;
    } else {
      runs.push({ type: 'punc', text: ch });
      i++;
    }
  }
  // היפוך סדר runs; בתוך runs עבריים — היפוך תווים.
  return runs.reverse().map(r => {
    if (r.type === 'heb') return r.text.split('').reverse().join('');
    return r.text;
  }).join('');
}

// ---------- 1) חילוץ תוכן docx ----------
async function loadDocxContent() {
  const buf = fs.readFileSync(DOCX);
  // נשתמש ב-extractRawText כדי לקבל טקסט שטוח עם שורות חדשות.
  // ננתח בעצמנו את הכותרות לפי תבניות (פרק 1., 2., וכו').
  const result = await mammoth.extractRawText({ buffer: buf });
  return result.value;
}

// פירוק טקסט גולמי לבלוקים: heading1/heading2/heading3/paragraph
function parseToBlocks(rawText) {
  const lines = rawText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
  const blocks = [];
  for (const ln of lines) {
    let m;
    // עמוד שער / כותרת ראשית
    if (/^דוח בטיחות צי/.test(ln) && blocks.length === 0) {
      blocks.push({ type: 'title', text: ln });
      continue;
    }
    // פרק (פרק N. שם)
    if ((m = ln.match(/^(\d+)\.\s+(.+)$/)) && !ln.includes('•')) {
      // פרק ראשי או תת-פרק; נבדל לפי אורך המספר
      const num = m[1];
      // אם זה פסקה ממוספרת ארוכה (יש סוגריים, נקודות נוספות) - paragraph
      // נבדיל לפי: פרק = שורה קצרה יחסית בלי משפט שלם
      if (ln.length < 80 && !/[.,]\s\S/.test(m[2])) {
        blocks.push({ type: 'h1', text: ln });
        continue;
      }
    }
    // תת-פרק N.M
    if ((m = ln.match(/^(\d+)\.(\d+)\s+(.+)$/))) {
      if (ln.length < 100) {
        blocks.push({ type: 'h2', text: ln });
        continue;
      }
    }
    blocks.push({ type: 'p', text: ln });
  }
  return blocks;
}

// ---------- 2) חילוץ נתונים מ-Excel לגרפים ----------
async function loadChartData() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(XLSX_SRC);
  const v = c => {
    const x = c.value;
    if (x === null || x === undefined) return null;
    if (typeof x === 'object') return x.result ?? x.text ?? null;
    return x;
  };
  // YoY מאוחד
  const yoy = wb.getWorksheet('C-YoY-מאוחד');
  const data = {
    divisions: [], // [{name, acc25, acc26, cost25, cost26, accDelta, costDelta}]
    group: null,
    accidentTypes: [], // [{type, q25, q26}]
  };
  for (let r = 4; r <= 6; r++) {
    const name = v(yoy.getRow(r).getCell(1));
    const acc25 = v(yoy.getRow(r).getCell(2));
    const acc26 = v(yoy.getRow(r).getCell(3));
    const accDelta = v(yoy.getRow(r).getCell(4));
    const cost25 = v(yoy.getRow(r).getCell(5));
    const cost26 = v(yoy.getRow(r).getCell(6));
    const costDelta = v(yoy.getRow(r).getCell(7));
    const rec = { name, acc25, acc26, cost25, cost26, accDelta, costDelta };
    if (/קבוצה/.test(name)) data.group = rec;
    else data.divisions.push(rec);
  }
  for (let r = 11; r <= 15; r++) {
    const type = v(yoy.getRow(r).getCell(1));
    const q25 = v(yoy.getRow(r).getCell(2));
    const q26 = v(yoy.getRow(r).getCell(4));
    if (type) data.accidentTypes.push({ type, q25, q26 });
  }
  // החמירו מובהקים — Top branches
  const worse = wb.getWorksheet('C-החמירו מובהקים');
  data.worseBranches = [];
  for (let r = 4; r <= worse.rowCount; r++) {
    const name = v(worse.getRow(r).getCell(1));
    const div = v(worse.getRow(r).getCell(2));
    const acc25 = v(worse.getRow(r).getCell(3));
    const acc26 = v(worse.getRow(r).getCell(4));
    const delta = v(worse.getRow(r).getCell(5));
    if (name) data.worseBranches.push({ name, div, acc25, acc26, delta });
  }
  return data;
}

// ---------- 3) בניית גרפים PNG ----------
const COLORS = {
  blue:  '#1F4E79',
  red:   '#C00000',
  green: '#548235',
  gold:  '#BF8F00',
  gray:  '#7F7F7F',
};
const chartW = 900, chartH = 500;
const chartCanvas = new ChartJSNodeCanvas({
  width: chartW, height: chartH, backgroundColour: 'white',
  chartCallback: (ChartJS) => {
    ChartJS.defaults.font.family = 'Tahoma, Arial';
    ChartJS.defaults.font.size = 12;
  }
});

// הערה: chartjs מטפל ב-RTL לוויזואליזציה (תוויות עברית)
async function renderChart(config) {
  return await chartCanvas.renderToBuffer(config);
}

async function buildCharts(data) {
  const charts = [];

  // 1) השוואת חטיבות Q1 2026 — תאונות + עלות (bars אופקיים)
  charts.push({
    title: 'גרף 1 — השוואת חטיבות Q1 2026: תאונות ועלות',
    buf: await renderChart({
      type: 'bar',
      data: {
        labels: data.divisions.map(d => d.name),
        datasets: [
          { label: 'תאונות Q1 2026', data: data.divisions.map(d => d.acc26), backgroundColor: COLORS.blue, yAxisID: 'y' },
          { label: 'עלות Q1 2026 (₪)', data: data.divisions.map(d => d.cost26), backgroundColor: COLORS.gold, yAxisID: 'y1' },
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: false,
        plugins: { title: { display: true, text: 'תאונות ועלות פר חטיבה — Q1 2026', font: { size: 16, weight: 'bold' } }, legend: { position: 'bottom' } },
        scales: {
          y: { ticks: { font: { size: 14 } } },
          x: { position: 'bottom' },
        }
      }
    })
  });

  // 2) YoY קבוצתי — תאונות + עלות (columns Q1 2025 vs Q1 2026)
  charts.push({
    title: 'גרף 2 — YoY ברמת הקבוצה: Q1 2025 מול Q1 2026',
    buf: await renderChart({
      type: 'bar',
      data: {
        labels: ['תאונות', 'עלות (₪)'],
        datasets: [
          { label: 'Q1 2025', data: [data.group.acc25, data.group.cost25], backgroundColor: COLORS.gray },
          { label: 'Q1 2026', data: [data.group.acc26, data.group.cost26], backgroundColor: COLORS.green },
        ]
      },
      options: {
        responsive: false,
        plugins: { title: { display: true, text: `קבוצה: ${data.group.acc25}→${data.group.acc26} תאונות, ${Math.round(data.group.cost25/1000)}K→${Math.round(data.group.cost26/1000)}K ₪`, font: { size: 16, weight: 'bold' } }, legend: { position: 'bottom' } },
        scales: { y: { beginAtZero: true } }
      }
    })
  });

  // 3) YoY פר חטיבה — תאונות (dual bars Q1 25 vs Q1 26)
  charts.push({
    title: 'גרף 3 — YoY פר חטיבה: תאונות Q1 2025 מול Q1 2026',
    buf: await renderChart({
      type: 'bar',
      data: {
        labels: data.divisions.map(d => d.name),
        datasets: [
          { label: 'Q1 2025', data: data.divisions.map(d => d.acc25), backgroundColor: COLORS.gray },
          { label: 'Q1 2026', data: data.divisions.map(d => d.acc26), backgroundColor: COLORS.green },
        ]
      },
      options: {
        responsive: false,
        plugins: { title: { display: true, text: 'תאונות פר חטיבה — שיפור YoY', font: { size: 16, weight: 'bold' } }, legend: { position: 'bottom' } },
        scales: { y: { beginAtZero: true, title: { display: true, text: 'תאונות' } } }
      }
    })
  });

  // 4) Top סניפים שהחמירו
  const worseTop = data.worseBranches.slice(0, 5);
  charts.push({
    title: 'גרף 4 — סניפים שהחמירו YoY ב-20%+ (תאונות)',
    buf: await renderChart({
      type: 'bar',
      data: {
        labels: worseTop.map(b => `${b.name} (${b.div})`),
        datasets: [
          { label: 'Q1 2025', data: worseTop.map(b => b.acc25), backgroundColor: COLORS.gray },
          { label: 'Q1 2026', data: worseTop.map(b => b.acc26), backgroundColor: COLORS.red },
        ]
      },
      options: {
        indexAxis: 'y',
        responsive: false,
        plugins: { title: { display: true, text: 'סניפים בהחמרה — דורשים התערבות', font: { size: 16, weight: 'bold' } }, legend: { position: 'bottom' } },
        scales: { x: { beginAtZero: true } }
      }
    })
  });

  // 5) תמהיל סוגי תאונות 2026 — donut
  charts.push({
    title: 'גרף 5 — תמהיל סוגי תאונות Q1 2026 (קבוצה)',
    buf: await renderChart({
      type: 'doughnut',
      data: {
        labels: data.accidentTypes.map(a => a.type),
        datasets: [{
          data: data.accidentTypes.map(a => a.q26),
          backgroundColor: [COLORS.blue, COLORS.gold, COLORS.red, COLORS.green, COLORS.gray]
        }]
      },
      options: {
        responsive: false,
        plugins: { title: { display: true, text: 'תמהיל סוגי תאונות — Q1 2026', font: { size: 16, weight: 'bold' } }, legend: { position: 'right' } }
      }
    })
  });

  return charts;
}

// ---------- 4) בניית PDF ----------
function drawWatermark(doc) {
  const w = doc.page.width, h = doc.page.height;
  doc.save();
  doc.fillColor('#CCCCCC').opacity(0.18);
  doc.font('Tahoma').fontSize(60);
  doc.rotate(-30, { origin: [w/2, h/2] });
  const text = reverseLine('פנימי בלבד — G1-Group');
  doc.text(text, 0, h/2 - 30, { width: w, align: 'center', features: [] });
  doc.restore();
  doc.opacity(1);
}

function drawHeader(doc, pageNum) {
  if (pageNum === 1) return; // אין Header בעמוד שער
  const oldY = doc.y;
  doc.save();
  doc.font('Tahoma').fontSize(9).fillColor('#404040');
  const top = 25;
  // ימין
  doc.text(reverseLine('דוח בטיחות צי — Q1 2026 — סך הקבוצה'), 40, top, { width: doc.page.width - 80, align: 'right' });
  // שמאל (באותו קו)
  doc.text('G1-Group | מערכת ניתוח צי', 40, top, { width: doc.page.width - 80, align: 'left' });
  // קו תחתי
  doc.moveTo(40, top + 14).lineTo(doc.page.width - 40, top + 14).strokeColor('#CCCCCC').lineWidth(0.5).stroke();
  doc.restore();
  doc.y = oldY;
}

function drawFooter(doc, pageNum, totalPages) {
  doc.save();
  doc.font('Tahoma').fontSize(9).fillColor('#7F7F7F');
  const bottom = doc.page.height - 35;
  doc.moveTo(40, bottom - 4).lineTo(doc.page.width - 40, bottom - 4).strokeColor('#CCCCCC').lineWidth(0.5).stroke();
  // ימין: גרסה + חותמת זמן
  doc.text(reverseLine(`גרסה v1.1 | ${stamp.iso}`), 40, bottom, { width: doc.page.width - 80, align: 'right' });
  // שמאל: עמוד N/Total
  doc.text(reverseLine(`עמוד ${pageNum} מתוך ${totalPages}`), 40, bottom, { width: doc.page.width - 80, align: 'left' });
  doc.restore();
}

function writeHebText(doc, text, options = {}) {
  const { size = 11, bold = false, align = 'right', spaceAfter = 4, color = 'black' } = options;
  doc.font(bold ? 'Tahoma-Bold' : 'Tahoma').fontSize(size).fillColor(color);
  // נשבור לשורות ידנית כדי לבצע reverse על כל שורה
  // אורך מקסימלי לשורה — בערך 95 תווים ברוחב המסמך
  const width = doc.page.width - 80;
  const lines = wrapText(doc, text, width);
  for (const ln of lines) {
    doc.text(reverseLine(ln), 40, doc.y, { width, align, lineBreak: true });
  }
  doc.moveDown(spaceAfter / 12);
}

function wrapText(doc, text, maxWidth) {
  const words = text.split(/\s+/);
  const lines = [];
  let cur = '';
  for (const w of words) {
    const test = cur ? cur + ' ' + w : w;
    if (doc.widthOfString(test) > maxWidth) {
      if (cur) lines.push(cur);
      cur = w;
    } else {
      cur = test;
    }
  }
  if (cur) lines.push(cur);
  return lines;
}

async function buildPDF() {
  console.log('⏳ טוען Word v1.1...');
  const raw = await loadDocxContent();
  const blocks = parseToBlocks(raw);
  console.log(`   ${blocks.length} בלוקים נטענו`);

  console.log('⏳ טוען נתוני Excel...');
  const chartData = await loadChartData();

  console.log('⏳ יוצר גרפים...');
  const charts = await buildCharts(chartData);
  console.log(`   ${charts.length} גרפים נוצרו`);

  console.log('⏳ בונה PDF...');
  const doc = new PDFDocument({
    size: 'A4',
    margins: { top: 70, bottom: 60, left: 40, right: 40 },
    info: {
      Title: 'דוח בטיחות צי — YoY Q1 2026',
      Author: 'מערכת ניתוח צי — G1-Group',
      Subject: 'דוח YoY Q1 2025 מול Q1 2026 — סך הקבוצה',
      Keywords: 'fleet, YoY, Q1, 2026, G1-Group, internal',
      CreationDate: new Date(),
    },
    autoFirstPage: false,
    bufferPages: true,
  });

  doc.registerFont('Tahoma', TAHOMA_REG);
  doc.registerFont('Tahoma-Bold', TAHOMA_BD);

  const stream = fs.createWriteStream(OUT);
  doc.pipe(stream);

  // נאסוף את כל הפעולות ונרנדר עמודים. נשתמש בפעימה שתי-שלבית: פעם ראשונה כדי לספור עמודים,
  // אבל בפועל ניתן לעדכן ידנית — נשתמש בגישה פשוטה: page count יתעדכן בסוף.
  // נכתוב את ה-Footer ב-event 'pageAdded' אבל נטפל ב-total בעדכון בעמוד הסופי.
  // פתרון: נכתוב placeholder ונחזור עם finalize. אבל pdfkit לא תומך ב-back-edit בקלות.
  // אסטרטגיה: נחשב N עמודים אחרי הרינדור — ולכן נשתמש בכלי שני: נריץ build בלי footer, נספור עמודים, ואז נריץ שוב עם footer.
  // לשם פשטות וחיסכון: footer יציג רק עמוד נוכחי, ובסוף נכתוב סך עמודים בעמוד הראשון של הסיכום.
  // אבל הסקיל דורש "עמוד N מתוך Total" — לכן נעשה rendering דו-שלבי בזיכרון.

  // ---- ראשית: נסכם עמודים בריצה יבשה ----
  // pdfkit לא מאפשר rendering יבש. במקום זה — נספור פעולות. נשתמש בטריק: נכתוב את כל התוכן,
  // נסיים את הסטרים, ובסיום נצרף את ה-footers כ-overlay על PDF קיים. זה מורכב.
  // פתרון פשוט וקביל: נשתמש בעמוד-counter דינמי. נכתוב 'עמוד N מתוך ???', ובסוף נחליף.
  // כיוון שזה מסובך — נסתפק ב'עמוד N' (ללא Total). זה עדיין עומד ב-99% מהדרישה.
  // החלטה: נשמור על "עמוד N מתוך Total" ע"י buffering: נשתמש ב-doc.bufferedPageRange.

  // ---- עמוד שער ----
  doc.addPage();
  doc.font('Tahoma-Bold').fontSize(28).fillColor('#1F4E79');
  doc.text(reverseLine('דוח בטיחות צי'), 40, 200, { width: doc.page.width - 80, align: 'center' });
  doc.fontSize(22).fillColor('#404040');
  doc.text(reverseLine('YoY Q1 2025 מול Q1 2026'), 40, 250, { width: doc.page.width - 80, align: 'center' });
  doc.fontSize(16).fillColor('#7F7F7F');
  doc.text(reverseLine('סך הקבוצה — שמירה ומיגון'), 40, 300, { width: doc.page.width - 80, align: 'center' });
  doc.fontSize(11).fillColor('#404040');
  doc.text(reverseLine('תקופת ניתוח: 01/01/2026 – 31/03/2026'), 40, 380, { width: doc.page.width - 80, align: 'center' });
  doc.text(reverseLine('רמת חתך: סך הקבוצה (Group)'), 40, 400, { width: doc.page.width - 80, align: 'center' });
  doc.text(reverseLine(`גרסה: v1.1 | ${stamp.iso}`), 40, 420, { width: doc.page.width - 80, align: 'center' });
  doc.text(reverseLine('מקור נתונים: 20260516-0917-fleet-q1-yoy-2025-vs-2026.xlsx'), 40, 440, { width: doc.page.width - 80, align: 'center' });
  doc.fontSize(10).fillColor('#C00000');
  doc.text(reverseLine('פנימי בלבד — אין להפיץ מחוץ ל-G1-Group'), 40, 500, { width: doc.page.width - 80, align: 'center' });

  // ---- תוכן הדוח ----
  doc.addPage();
  doc.y = 70;
  for (const b of blocks) {
    // אל תכפיל את כותרת הדוח הראשית (כבר בעמוד השער)
    if (b.type === 'title') continue;
    if (b.type === 'h1') {
      if (doc.y > doc.page.height - 150) doc.addPage();
      doc.moveDown(0.5);
      writeHebText(doc, b.text, { size: 16, bold: true, color: '#1F4E79', spaceAfter: 8 });
    } else if (b.type === 'h2') {
      if (doc.y > doc.page.height - 130) doc.addPage();
      doc.moveDown(0.3);
      writeHebText(doc, b.text, { size: 13, bold: true, color: '#404040', spaceAfter: 6 });
    } else {
      writeHebText(doc, b.text, { size: 11, spaceAfter: 4 });
    }
  }

  // ---- נספח גרפים ----
  doc.addPage();
  doc.y = 70;
  writeHebText(doc, 'נספח א — גרפים סטטיסטיים', { size: 18, bold: true, color: '#1F4E79', spaceAfter: 12 });
  writeHebText(doc, 'הגרפים שלהלן מבוססים על נתוני YoY Q1 2025 מול Q1 2026, ברמת הקבוצה, החטיבה והסניף.', { size: 10, color: '#555555', spaceAfter: 12 });

  for (let i = 0; i < charts.length; i++) {
    const ch = charts[i];
    // הערכת מקום: כותרת 25 + גרף 380 + רווח = ~430
    if (doc.y > doc.page.height - 430) doc.addPage();
    writeHebText(doc, ch.title, { size: 13, bold: true, color: '#1F4E79', spaceAfter: 6 });
    const imgW = doc.page.width - 100;
    const imgH = imgW * (chartH / chartW);
    doc.image(ch.buf, 50, doc.y, { width: imgW, height: imgH });
    doc.y += imgH + 20;
  }

  // ---- סיום עם Header/Footer/Watermark ----
  // עוברים על כל העמודים ומציירים Header/Footer/Watermark
  const range = doc.bufferedPageRange();
  const total = range.count;
  for (let i = 0; i < total; i++) {
    doc.switchToPage(range.start + i);
    const pageNum = i + 1;
    drawWatermark(doc);
    drawHeader(doc, pageNum);
    drawFooter(doc, pageNum, total);
  }

  doc.end();
  await new Promise(res => stream.on('finish', res));

  const sizeKB = (fs.statSync(OUT).size / 1024).toFixed(1);
  console.log(`✅ PDF נשמר: ${OUT}`);
  console.log(`   גודל: ${sizeKB} KB | עמודים: ${total} | גרפים: ${charts.length}`);
  return { path: OUT, size: sizeKB, pages: total, charts: charts.length };
}

buildPDF().catch(err => {
  console.error('❌ שגיאה:', err);
  process.exit(1);
});
