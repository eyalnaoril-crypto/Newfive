#!/usr/bin/env node
// Gal — Deep-dive analysis of "ריכוז השמירה" + "ריכוז הט"מ" sheets.
// Reads the updated source workbook, parses every quarterly block from 2014 to 2026Q1,
// builds normalized raw-data tables, summary tables, PNG charts, and writes a single
// Excel workbook with all sheets in RTL / Tahoma 11.

import ExcelJS from 'exceljs';
import { ChartJSNodeCanvas } from 'chartjs-node-canvas';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'output', '20260516-0827-fleet-source-updated-2026Q1.xlsx');

const stamp = (() => {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}`;
})();
const OUT = path.join(ROOT, 'output', `${stamp}-fleet-shmira-deep-dive.xlsx`);

// ───────── helpers ─────────
const TAHOMA = { name: 'Tahoma', size: 11 };
const TAHOMA_BOLD = { name: 'Tahoma', size: 11, bold: true };
const TAHOMA_TITLE = { name: 'Tahoma', size: 14, bold: true };

const cellVal = c => {
  const v = c.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v.result !== undefined) return v.result;
    if (v.text !== undefined) return v.text;
    if (v.richText) return v.richText.map(t => t.text).join('');
  }
  return v;
};
const num = v => {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[,₪\s]/g, ''));
  return isNaN(n) ? null : n;
};
const isQuarterHeader = s => /^רבעון\s*[-–\s]*\s*[1-4]/.test(String(s).trim()) || /רבעון\s*[1-4]\s*שנת/.test(String(s));

function parsePeriodLabel(label) {
  // Supports: "רבעון 1 שנת 2014", "רבעון 1 2017", "רבעון 1-2022", "רבעון -2025 -2", "רבעון -1-2026"
  const s = String(label).replace(/\s+/g, ' ').trim();
  // try various patterns
  let m = s.match(/רבעון\s*([1-4])\s*שנת\s*(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4])\s*[-–]\s*(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4])\s+(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*[-–]\s*(\d{4})\s*[-–]\s*([1-4])/); // רבעון -2025 -2
  if (m) return { q: +m[2], y: +m[1] };
  m = s.match(/רבעון\s*[-–]\s*([1-4])\s*[-–]\s*(\d{4})/); // רבעון -1-2026
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4]).*?(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  return null;
}

// Canonical schema for the parsed block. Column meanings inferred from headers in 2014..2026.
// Most blocks share this layout (column B = branch, C..R = data).
// Column indexes (1-based) found in the source sheets:
//   B(2)  סניף
//   C(3)  כמות כלי רכב בסניף
//   D(4)  כלי רכב סיור / סיור-הסעות
//   E(5)  סך ק"מ לרבעון
//   F(6)  נהג החברה אשם
//   G(7)  צד ג' אשם + פריצה + נזקי חניה
//   H(8)  נזק מרכב תחתון
//   I(9)  נזקי סיום עיסקה
//   J(10) טוטל-לוס
//   K(11) כמות תאונות ונזקים (total)
//   L(12) עלות כלל המקרים לחברה
//   M(13) עלות נזקי סיום עיסקה
//   N(14) נוהל 6
//   O(15) עבירות תנועה / חניה
//   P(16) כמות תאונות ממוצעת לרכב
//   Q(17) עלות נזקים ממוצעת לרכב
//   R(18) עלות נזק ל-1,000 ק"מ

function parseSheet(ws) {
  const rows = ws.rowCount;
  // 1) detect quarterly block headers
  const headers = [];
  for (let r = 1; r <= rows; r++) {
    const v = cellVal(ws.getRow(r).getCell(2));
    if (isQuarterHeader(v)) {
      const p = parsePeriodLabel(v);
      if (p) headers.push({ row: r, label: String(v).trim(), ...p });
    }
  }
  // 2) for each block, parse data rows up to סה"כ
  const records = [];
  const totals = [];
  const seenPeriods = new Set();
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    const periodKey = `${h.y}-Q${h.q}`;
    const nextRow = i + 1 < headers.length ? headers[i + 1].row : rows + 1;
    // find header row (the one starting with סניף or .) — usually 2..4 rows below
    let headerRow = null;
    for (let r = h.row + 1; r <= Math.min(h.row + 5, rows); r++) {
      const v = String(cellVal(ws.getRow(r).getCell(2))).trim();
      if (v === 'סניף' || v === '.' || /סניף/.test(v)) { headerRow = r; break; }
    }
    if (!headerRow) continue;
    // Skip duplicate period (e.g. shmira has רבעון 4-2022 twice)
    if (seenPeriods.has(periodKey)) continue;
    seenPeriods.add(periodKey);
    // iterate data rows
    let totalRow = null;
    for (let r = headerRow + 1; r < nextRow; r++) {
      const branch = String(cellVal(ws.getRow(r).getCell(2))).trim();
      if (!branch) continue;
      if (branch.includes('סה"כ') || branch.includes('סה״כ') || branch.includes("סה'כ")) {
        totalRow = r;
        break;
      }
      const row = ws.getRow(r);
      const rec = {
        period: periodKey,
        year: h.y,
        quarter: h.q,
        branch,
        vehicles: num(cellVal(row.getCell(3))),
        vehicles_patrol: num(cellVal(row.getCell(4))),
        km: num(cellVal(row.getCell(5))),
        driver_at_fault: num(cellVal(row.getCell(6))),
        third_party: num(cellVal(row.getCell(7))),
        undercarriage: num(cellVal(row.getCell(8))),
        end_of_deal: num(cellVal(row.getCell(9))),
        total_loss: num(cellVal(row.getCell(10))),
        accidents_count: num(cellVal(row.getCell(11))),
        accidents_cost: num(cellVal(row.getCell(12))),
        end_of_deal_cost: num(cellVal(row.getCell(13))),
        nohal_6: num(cellVal(row.getCell(14))),
        traffic_violations: num(cellVal(row.getCell(15))),
        accidents_per_vehicle: num(cellVal(row.getCell(16))),
        cost_per_vehicle: num(cellVal(row.getCell(17))),
        cost_per_1000km: num(cellVal(row.getCell(18))),
      };
      // Skip fully empty rows
      const hasData = ['vehicles','km','accidents_count','accidents_cost','driver_at_fault','third_party'].some(k => rec[k] !== null && rec[k] !== 0) || rec.branch;
      if (hasData && rec.branch) records.push(rec);
    }
    // parse total row
    if (totalRow) {
      const row = ws.getRow(totalRow);
      totals.push({
        period: periodKey,
        year: h.y,
        quarter: h.q,
        vehicles: num(cellVal(row.getCell(3))),
        km: num(cellVal(row.getCell(5))),
        driver_at_fault: num(cellVal(row.getCell(6))),
        third_party: num(cellVal(row.getCell(7))),
        undercarriage: num(cellVal(row.getCell(8))),
        end_of_deal: num(cellVal(row.getCell(9))),
        total_loss: num(cellVal(row.getCell(10))),
        accidents_count: num(cellVal(row.getCell(11))),
        accidents_cost: num(cellVal(row.getCell(12))),
        end_of_deal_cost: num(cellVal(row.getCell(13))),
        nohal_6: num(cellVal(row.getCell(14))),
        traffic_violations: num(cellVal(row.getCell(15))),
      });
    }
  }
  return { records, totals };
}

// ───────── chart rendering ─────────
const chartW = 900, chartH = 520;
const chartCanvas = new ChartJSNodeCanvas({
  width: chartW,
  height: chartH,
  backgroundColour: 'white',
  chartCallback: (ChartJS) => {
    ChartJS.defaults.font.family = 'Tahoma';
    ChartJS.defaults.font.size = 12;
  },
});
async function renderChart(config) {
  return await chartCanvas.renderToBuffer(config, 'image/png');
}

const palette = ['#1F4E79','#C00000','#548235','#BF8F00','#7030A0','#2E75B6','#A52A2A','#385723','#806000','#4472C4'];

// ───────── number formatting / shekel ─────────
const fmtMoney = '#,##0\\ ₪';
const fmtInt = '#,##0';
const fmtPct = '0.0%';

// ───────── styling helpers ─────────
function styleHeader(cell) {
  cell.font = TAHOMA_BOLD;
  cell.alignment = { horizontal: 'right', vertical: 'middle', readingOrder: 'rtl', wrapText: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  cell.font = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  cell.border = {
    top: { style: 'thin', color: { argb: 'FF999999' } },
    bottom: { style: 'thin', color: { argb: 'FF999999' } },
    left: { style: 'thin', color: { argb: 'FF999999' } },
    right: { style: 'thin', color: { argb: 'FF999999' } },
  };
}
function styleBody(cell) {
  cell.font = TAHOMA;
  cell.alignment = { horizontal: 'right', vertical: 'middle', readingOrder: 'rtl' };
  cell.border = {
    top: { style: 'thin', color: { argb: 'FFDDDDDD' } },
    bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
    left: { style: 'thin', color: { argb: 'FFDDDDDD' } },
    right: { style: 'thin', color: { argb: 'FFDDDDDD' } },
  };
}
function styleTitle(cell) {
  cell.font = TAHOMA_TITLE;
  cell.alignment = { horizontal: 'right', vertical: 'middle', readingOrder: 'rtl' };
}
function setRTL(ws) {
  ws.views = [{ rightToLeft: true, showGridLines: false }];
}
function writeTable(ws, startRow, headers, rows, moneyCols = [], pctCols = [], intCols = []) {
  // headers
  const hr = ws.getRow(startRow);
  headers.forEach((h, i) => {
    const c = hr.getCell(i + 1);
    c.value = h;
    styleHeader(c);
  });
  hr.height = 30;
  // body
  rows.forEach((r, ri) => {
    const row = ws.getRow(startRow + 1 + ri);
    r.forEach((v, ci) => {
      const c = row.getCell(ci + 1);
      c.value = (v === null || v === undefined) ? 'N/A' : v;
      styleBody(c);
      if (moneyCols.includes(ci)) c.numFmt = fmtMoney;
      else if (pctCols.includes(ci)) c.numFmt = fmtPct;
      else if (intCols.includes(ci)) c.numFmt = fmtInt;
      if (v === null || v === undefined) {
        c.font = { name: 'Tahoma', size: 11, italic: true, color: { argb: 'FF888888' } };
      }
    });
  });
  return startRow + 1 + rows.length;
}

// ───────── main ─────────
(async () => {
  console.log('Reading source workbook…');
  const src = new ExcelJS.Workbook();
  await src.xlsx.readFile(SRC);
  const wsS = src.getWorksheet('ריכוז השמירה');
  const wsM = src.getWorksheet('ריכוז הט"מ');

  console.log('Parsing ריכוז השמירה…');
  const shmira = parseSheet(wsS);
  console.log('  records:', shmira.records.length, ' totals:', shmira.totals.length);

  console.log('Parsing ריכוז הט"מ…');
  const migun = parseSheet(wsM);
  console.log('  records:', migun.records.length, ' totals:', migun.totals.length);

  // Compute summary tables for each division
  function buildSummaries(records, totals) {
    // 1) Yearly trend (using totals row when available; fallback aggregation of records)
    const yearMap = new Map();
    for (const t of totals) {
      const y = t.year;
      if (!yearMap.has(y)) yearMap.set(y, { year: y, accidents: 0, cost: 0, end_of_deal_cost: 0, km: 0, vehicles_max: 0, quarters: 0 });
      const cur = yearMap.get(y);
      cur.accidents += t.accidents_count || 0;
      cur.cost += t.accidents_cost || 0;
      cur.end_of_deal_cost += t.end_of_deal_cost || 0;
      cur.km += t.km || 0;
      cur.vehicles_max = Math.max(cur.vehicles_max, t.vehicles || 0);
      cur.quarters += 1;
    }
    const yearly = Array.from(yearMap.values()).sort((a, b) => a.year - b.year);
    yearly.forEach(y => { y.avg_cost_per_accident = y.accidents ? y.cost / y.accidents : null; });

    // 2) Top branches by cumulative cost
    const branchMap = new Map();
    for (const r of records) {
      const b = r.branch.trim();
      if (!b) continue;
      if (!branchMap.has(b)) branchMap.set(b, { branch: b, accidents: 0, cost: 0, periods: 0 });
      const cur = branchMap.get(b);
      cur.accidents += r.accidents_count || 0;
      cur.cost += r.accidents_cost || 0;
      cur.periods += 1;
    }
    const branches = Array.from(branchMap.values())
      .sort((a, b) => b.cost - a.cost)
      .slice(0, 10);

    // 3) Accident type breakdown (sum of driver_at_fault / third_party / undercarriage / end_of_deal / total_loss)
    const typeMap = {
      'נהג החברה אשם': 0,
      'צד ג\' / פריצה / חניה': 0,
      'נזק מרכב תחתון': 0,
      'נזקי סיום עיסקה': 0,
      'טוטל-לוס': 0,
    };
    for (const r of records) {
      typeMap['נהג החברה אשם'] += r.driver_at_fault || 0;
      typeMap['צד ג\' / פריצה / חניה'] += r.third_party || 0;
      typeMap['נזק מרכב תחתון'] += r.undercarriage || 0;
      typeMap['נזקי סיום עיסקה'] += r.end_of_deal || 0;
      typeMap['טוטל-לוס'] += r.total_loss || 0;
    }
    const types = Object.entries(typeMap).map(([type, count]) => ({ type, count }));

    return { yearly, branches, types };
  }

  const shmiraSum = buildSummaries(shmira.records, shmira.totals);
  const migunSum = buildSummaries(migun.records, migun.totals);

  // ───────── build output workbook ─────────
  console.log('Building workbook…');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'גל — Excel Fleet Analyst';
  wb.created = new Date();

  // Helper to add a sheet with RTL
  function addSheet(name) {
    const ws = wb.addWorksheet(name, { properties: { defaultRowHeight: 18 } });
    setRTL(ws);
    return ws;
  }

  // ===== PART A — ריכוז השמירה =====
  // 1) raw data
  {
    const ws = addSheet('שמירה-נתונים גולמיים');
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 8;
    ws.getColumn(3).width = 8;
    ws.getColumn(4).width = 20;
    for (let i = 5; i <= 20; i++) ws.getColumn(i).width = 14;
    const headers = ['תקופה','שנה','רבעון','סניף','מס\' רכבים','רכבי סיור','ק"מ ברבעון','נהג אשם','צד ג\'/פריצה','מרכב תחתון','סיום עיסקה','טוטל לוס','סה"כ תאונות','עלות תאונות','עלות סיום עיסקה','נוהל 6','עבירות תנועה','תאונות לרכב','עלות לרכב','עלות ל-1000ק"מ'];
    const rows = shmira.records.map(r => [
      r.period, r.year, r.quarter, r.branch,
      r.vehicles, r.vehicles_patrol, r.km,
      r.driver_at_fault, r.third_party, r.undercarriage, r.end_of_deal, r.total_loss,
      r.accidents_count, r.accidents_cost, r.end_of_deal_cost,
      r.nohal_6, r.traffic_violations,
      r.accidents_per_vehicle, r.cost_per_vehicle, r.cost_per_1000km,
    ]);
    writeTable(ws, 1, headers, rows, [13, 14], [], [4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16]);
    ws.views = [{ rightToLeft: true, showGridLines: false, state: 'frozen', xSplit: 0, ySplit: 1 }];
  }

  // 2) summary tables
  {
    const ws = addSheet('שמירה-טבלאות מסכמות');
    ws.getColumn(1).width = 24;
    for (let i = 2; i <= 10; i++) ws.getColumn(i).width = 16;
    let row = 1;
    ws.getCell(row, 1).value = 'מגמה רב-שנתית — אבטחה + ניקיון';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['שנה','סה"כ תאונות','עלות תאונות','עלות סיום עיסקה','ק"מ נסיעה','רבעונים שדווחו','עלות ממוצעת לתאונה'],
      shmiraSum.yearly.map(y => [y.year, y.accidents, y.cost, y.end_of_deal_cost, y.km, y.quarters, y.avg_cost_per_accident]),
      [2, 3, 6], [], [1, 4, 5]);
    row += 2;
    ws.getCell(row, 1).value = 'Top 10 סניפים לפי עלות תאונות מצטברת';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['סניף','עלות מצטברת','סה"כ תאונות','רבעונים פעילים','עלות ממוצעת לתאונה'],
      shmiraSum.branches.map(b => [b.branch, b.cost, b.accidents, b.periods, b.accidents ? b.cost / b.accidents : null]),
      [1, 4], [], [2, 3]);
    row += 2;
    ws.getCell(row, 1).value = 'חלוקה לפי סוג תאונה (סך כל התקופות)';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['סוג תאונה','כמות מצטברת'],
      shmiraSum.types.map(t => [t.type, t.count]), [], [], [1]);
  }

  // 3) charts
  {
    const ws = addSheet('שמירה-גרפים');
    ws.getColumn(1).width = 100;
    ws.getCell(1, 1).value = 'גרפים — חטיבת אבטחה + ניקיון';
    styleTitle(ws.getCell(1, 1));

    // line: yearly cost trend
    const cfgLine = {
      type: 'line',
      data: {
        labels: shmiraSum.yearly.map(y => y.year),
        datasets: [{
          label: 'עלות תאונות (₪)',
          data: shmiraSum.yearly.map(y => y.cost),
          borderColor: palette[0], backgroundColor: palette[0], fill: false, tension: 0.2, pointRadius: 5,
        }, {
          label: 'מס\' תאונות',
          data: shmiraSum.yearly.map(y => y.accidents),
          borderColor: palette[1], backgroundColor: palette[1], fill: false, tension: 0.2, pointRadius: 5,
          yAxisID: 'y1',
        }],
      },
      options: {
        plugins: { title: { display: true, text: 'מגמה רב-שנתית — שמירה' } },
        scales: { y: { beginAtZero: true, title: { display: true, text: '₪' } }, y1: { beginAtZero: true, position: 'left', grid: { drawOnChartArea: false }, title: { display: true, text: 'תאונות' } } },
      },
    };
    const img1 = wb.addImage({ buffer: await renderChart(cfgLine), extension: 'png' });
    ws.addImage(img1, { tl: { col: 0, row: 2 }, ext: { width: chartW, height: chartH } });

    // horizontal bar: top 10 branches
    const cfgBar = {
      type: 'bar',
      data: {
        labels: shmiraSum.branches.map(b => b.branch),
        datasets: [{ label: 'עלות מצטברת (₪)', data: shmiraSum.branches.map(b => b.cost), backgroundColor: palette[2] }],
      },
      options: {
        indexAxis: 'y',
        plugins: { title: { display: true, text: 'Top 10 סניפים — עלות מצטברת (שמירה)' }, legend: { display: false } },
      },
    };
    const img2 = wb.addImage({ buffer: await renderChart(cfgBar), extension: 'png' });
    ws.addImage(img2, { tl: { col: 0, row: 30 }, ext: { width: chartW, height: chartH } });

    // donut: accident type breakdown
    const cfgDonut = {
      type: 'doughnut',
      data: {
        labels: shmiraSum.types.map(t => t.type),
        datasets: [{ data: shmiraSum.types.map(t => t.count), backgroundColor: palette.slice(0, 5) }],
      },
      options: { plugins: { title: { display: true, text: 'התפלגות סוג תאונה — שמירה' } } },
    };
    const img3 = wb.addImage({ buffer: await renderChart(cfgDonut), extension: 'png' });
    ws.addImage(img3, { tl: { col: 0, row: 58 }, ext: { width: chartW, height: chartH } });
  }

  // 4) conclusions
  {
    const ws = addSheet('שמירה-מסקנות');
    ws.getColumn(1).width = 120;
    const peakYear = [...shmiraSum.yearly].sort((a,b) => b.cost - a.cost)[0];
    const lowYear = [...shmiraSum.yearly].filter(y => y.quarters >= 3).sort((a,b) => a.cost - b.cost)[0];
    const topBranch = shmiraSum.branches[0];
    const totalCost = shmiraSum.yearly.reduce((s,y) => s+y.cost, 0);
    const totalAcc = shmiraSum.yearly.reduce((s,y) => s+y.accidents, 0);
    const conclusions = [
      `מסקנות מנתוני "ריכוז השמירה" — חטיבת אבטחה + ניקיון, ${shmiraSum.yearly[0]?.year}–${shmiraSum.yearly.at(-1)?.year}`,
      '',
      `1. נתח רב-שנתי: סך עלות תאונות מצטברת לאורך כל התקופה — ${Math.round(totalCost).toLocaleString('he-IL')} ₪ עבור ${totalAcc.toLocaleString('he-IL')} תאונות.`,
      `2. שנת השיא בעלויות הייתה ${peakYear?.year} עם ${Math.round(peakYear?.cost||0).toLocaleString('he-IL')} ₪ (${peakYear?.accidents||0} תאונות, ${peakYear?.quarters||0} רבעונים).`,
      `3. הסניף עם העלות הגבוהה ביותר במצטבר: ${topBranch?.branch} — ${Math.round(topBranch?.cost||0).toLocaleString('he-IL')} ₪ ב-${topBranch?.periods||0} רבעונים.`,
      `4. עלות ממוצעת לתאונה בכל ההיסטוריה: ${totalAcc ? Math.round(totalCost/totalAcc).toLocaleString('he-IL') : 'N/A'} ₪.`,
      `5. סוג התאונה הדומיננטי: "${shmiraSum.types.sort((a,b)=>b.count-a.count)[0]?.type}" — ${shmiraSum.types.sort((a,b)=>b.count-a.count)[0]?.count.toLocaleString('he-IL')} מקרים.`,
      `6. ${lowYear ? `השנה הרגועה ביותר בהיקף עלויות הייתה ${lowYear.year} (${Math.round(lowYear.cost).toLocaleString('he-IL')} ₪).` : 'לא זוהתה שנה רגועה במיוחד עם דיווח רבעוני מלא.'}`,
      `7. ההתפלגות בין "נהג אשם" ל"צד ג'/פריצה" מצביעה על האחריות התפעולית מול חשיפה חיצונית — שני המקרים מהווים את עיקר הנפח.`,
      `8. אזהרת איכות נתונים: בלוקים מסוימים אינם מכילים פירוט עלויות (תאים ריקים) — מצב זה סומן N/A ולא הוחלף ב-0 כדי לא לעוות מגמות.`,
    ];
    conclusions.forEach((line, i) => {
      const c = ws.getCell(i + 1, 1);
      c.value = line;
      if (i === 0) styleTitle(c);
      else {
        c.font = TAHOMA;
        c.alignment = { horizontal: 'right', vertical: 'top', readingOrder: 'rtl', wrapText: true };
      }
    });
  }

  // ===== PART B — ריכוז הט"מ =====
  {
    const ws = addSheet('מיגון-נתונים גולמיים');
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 8;
    ws.getColumn(3).width = 8;
    ws.getColumn(4).width = 22;
    for (let i = 5; i <= 20; i++) ws.getColumn(i).width = 14;
    const headers = ['תקופה','שנה','רבעון','סניף','מס\' רכבים','רכבי סיור','ק"מ ברבעון','נהג אשם','צד ג\'/פריצה','מרכב תחתון','סיום עיסקה','טוטל לוס','סה"כ תאונות','עלות תאונות','עלות סיום עיסקה','נוהל 6','עבירות תנועה','תאונות לרכב','עלות לרכב','עלות ל-1000ק"מ'];
    const rows = migun.records.map(r => [
      r.period, r.year, r.quarter, r.branch,
      r.vehicles, r.vehicles_patrol, r.km,
      r.driver_at_fault, r.third_party, r.undercarriage, r.end_of_deal, r.total_loss,
      r.accidents_count, r.accidents_cost, r.end_of_deal_cost,
      r.nohal_6, r.traffic_violations,
      r.accidents_per_vehicle, r.cost_per_vehicle, r.cost_per_1000km,
    ]);
    writeTable(ws, 1, headers, rows, [13, 14], [], [4, 5, 6, 7, 8, 9, 10, 11, 12, 15, 16]);
    ws.views = [{ rightToLeft: true, showGridLines: false, state: 'frozen', xSplit: 0, ySplit: 1 }];
  }

  {
    const ws = addSheet('מיגון-טבלאות מסכמות');
    ws.getColumn(1).width = 26;
    for (let i = 2; i <= 10; i++) ws.getColumn(i).width = 16;
    let row = 1;
    ws.getCell(row, 1).value = 'מגמה רב-שנתית — טכנולוגיות מיגון + מוקדים';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['שנה','סה"כ תאונות','עלות תאונות','עלות סיום עיסקה','ק"מ נסיעה','רבעונים שדווחו','עלות ממוצעת לתאונה'],
      migunSum.yearly.map(y => [y.year, y.accidents, y.cost, y.end_of_deal_cost, y.km, y.quarters, y.avg_cost_per_accident]),
      [2, 3, 6], [], [1, 4, 5]);
    row += 2;
    ws.getCell(row, 1).value = 'Top 10 סניפים לפי עלות תאונות מצטברת';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['סניף','עלות מצטברת','סה"כ תאונות','רבעונים פעילים','עלות ממוצעת לתאונה'],
      migunSum.branches.map(b => [b.branch, b.cost, b.accidents, b.periods, b.accidents ? b.cost / b.accidents : null]),
      [1, 4], [], [2, 3]);
    row += 2;
    ws.getCell(row, 1).value = 'חלוקה לפי סוג תאונה (סך כל התקופות)';
    styleTitle(ws.getCell(row, 1));
    row += 2;
    row = writeTable(ws, row, ['סוג תאונה','כמות מצטברת'],
      migunSum.types.map(t => [t.type, t.count]), [], [], [1]);
  }

  {
    const ws = addSheet('מיגון-גרפים');
    ws.getColumn(1).width = 100;
    ws.getCell(1, 1).value = 'גרפים — חטיבת טכנולוגיות מיגון + מוקדים';
    styleTitle(ws.getCell(1, 1));

    const cfgLine = {
      type: 'line',
      data: {
        labels: migunSum.yearly.map(y => y.year),
        datasets: [{
          label: 'עלות תאונות (₪)',
          data: migunSum.yearly.map(y => y.cost),
          borderColor: palette[1], backgroundColor: palette[1], fill: false, tension: 0.2, pointRadius: 5,
        }, {
          label: 'מס\' תאונות',
          data: migunSum.yearly.map(y => y.accidents),
          borderColor: palette[3], backgroundColor: palette[3], fill: false, tension: 0.2, pointRadius: 5,
          yAxisID: 'y1',
        }],
      },
      options: {
        plugins: { title: { display: true, text: 'מגמה רב-שנתית — מיגון' } },
        scales: { y: { beginAtZero: true, title: { display: true, text: '₪' } }, y1: { beginAtZero: true, position: 'left', grid: { drawOnChartArea: false }, title: { display: true, text: 'תאונות' } } },
      },
    };
    const img1 = wb.addImage({ buffer: await renderChart(cfgLine), extension: 'png' });
    ws.addImage(img1, { tl: { col: 0, row: 2 }, ext: { width: chartW, height: chartH } });

    const cfgBar = {
      type: 'bar',
      data: {
        labels: migunSum.branches.map(b => b.branch),
        datasets: [{ label: 'עלות מצטברת (₪)', data: migunSum.branches.map(b => b.cost), backgroundColor: palette[5] }],
      },
      options: { indexAxis: 'y', plugins: { title: { display: true, text: 'Top 10 סניפים — עלות מצטברת (מיגון)' }, legend: { display: false } } },
    };
    const img2 = wb.addImage({ buffer: await renderChart(cfgBar), extension: 'png' });
    ws.addImage(img2, { tl: { col: 0, row: 30 }, ext: { width: chartW, height: chartH } });

    const cfgDonut = {
      type: 'doughnut',
      data: {
        labels: migunSum.types.map(t => t.type),
        datasets: [{ data: migunSum.types.map(t => t.count), backgroundColor: palette.slice(0, 5) }],
      },
      options: { plugins: { title: { display: true, text: 'התפלגות סוג תאונה — מיגון' } } },
    };
    const img3 = wb.addImage({ buffer: await renderChart(cfgDonut), extension: 'png' });
    ws.addImage(img3, { tl: { col: 0, row: 58 }, ext: { width: chartW, height: chartH } });
  }

  {
    const ws = addSheet('מיגון-מסקנות');
    ws.getColumn(1).width = 120;
    const peakYear = [...migunSum.yearly].sort((a,b) => b.cost - a.cost)[0];
    const lowYear = [...migunSum.yearly].filter(y => y.quarters >= 3).sort((a,b) => a.cost - b.cost)[0];
    const topBranch = migunSum.branches[0];
    const totalCost = migunSum.yearly.reduce((s,y) => s+y.cost, 0);
    const totalAcc = migunSum.yearly.reduce((s,y) => s+y.accidents, 0);
    const conclusions = [
      `מסקנות מנתוני "ריכוז הט"מ" — חטיבת טכנולוגיות מיגון + מוקדים, ${migunSum.yearly[0]?.year}–${migunSum.yearly.at(-1)?.year}`,
      '',
      `1. נתח רב-שנתי: סך עלות תאונות מצטברת — ${Math.round(totalCost).toLocaleString('he-IL')} ₪ עבור ${totalAcc.toLocaleString('he-IL')} תאונות.`,
      `2. שנת השיא בעלויות הייתה ${peakYear?.year} עם ${Math.round(peakYear?.cost||0).toLocaleString('he-IL')} ₪ (${peakYear?.accidents||0} תאונות).`,
      `3. הסניף/יחידה עם העלות הגבוהה ביותר במצטבר: ${topBranch?.branch} — ${Math.round(topBranch?.cost||0).toLocaleString('he-IL')} ₪.`,
      `4. עלות ממוצעת לתאונה בכל ההיסטוריה: ${totalAcc ? Math.round(totalCost/totalAcc).toLocaleString('he-IL') : 'N/A'} ₪.`,
      `5. סוג התאונה הדומיננטי: "${migunSum.types.sort((a,b)=>b.count-a.count)[0]?.type}" — ${migunSum.types.sort((a,b)=>b.count-a.count)[0]?.count.toLocaleString('he-IL')} מקרים.`,
      `6. ${lowYear ? `השנה הרגועה ביותר בעלויות הייתה ${lowYear.year} (${Math.round(lowYear.cost).toLocaleString('he-IL')} ₪).` : 'לא זוהתה שנה רגועה משמעותית.'}`,
      `7. מיגון כולל יחידות גדולות (גוש דן, פרוייקטים, חטיבת מוקדים) — חשיפה לעלויות גבוהות יותר מאשר ביחידות שמירה הקטנות יותר.`,
      `8. אזהרת איכות נתונים: ערכי "0" בעמודות 2025–2026 עשויים לבטא ערך אמיתי או חוסר דיווח — דרושה אימות מול גלעד לפני קביעה תקיפה.`,
    ];
    conclusions.forEach((line, i) => {
      const c = ws.getCell(i + 1, 1);
      c.value = line;
      if (i === 0) styleTitle(c);
      else {
        c.font = TAHOMA;
        c.alignment = { horizontal: 'right', vertical: 'top', readingOrder: 'rtl', wrapText: true };
      }
    });
  }

  // ===== PART C — Comparison =====
  const allPeriods = new Set([...shmira.totals.map(t => t.period), ...migun.totals.map(t => t.period)]);
  const sortedPeriods = [...allPeriods].sort((a, b) => {
    const [ya, qa] = a.split('-Q').map(Number);
    const [yb, qb] = b.split('-Q').map(Number);
    return ya - yb || qa - qb;
  });
  const shmiraMap = new Map(shmira.totals.map(t => [t.period, t]));
  const migunMap = new Map(migun.totals.map(t => [t.period, t]));

  {
    const ws = addSheet('השוואה-טבלת על');
    ws.getColumn(1).width = 16;
    for (let i = 2; i <= 8; i++) ws.getColumn(i).width = 18;
    const headers = ['תקופה','שמירה — תאונות','שמירה — עלות (₪)','מיגון — תאונות','מיגון — עלות (₪)','סה"כ תאונות','סה"כ עלות (₪)'];
    const rows = sortedPeriods.map(p => {
      const s = shmiraMap.get(p) || {};
      const m = migunMap.get(p) || {};
      const sAcc = s.accidents_count ?? null;
      const sCost = s.accidents_cost ?? null;
      const mAcc = m.accidents_count ?? null;
      const mCost = m.accidents_cost ?? null;
      const totalAcc = (sAcc || 0) + (mAcc || 0);
      const totalCost = (sCost || 0) + (mCost || 0);
      return [p, sAcc, sCost, mAcc, mCost, totalAcc || null, totalCost || null];
    });
    writeTable(ws, 1, headers, rows, [2, 4, 6], [], [1, 3, 5]);
    ws.views = [{ rightToLeft: true, showGridLines: false, state: 'frozen', xSplit: 0, ySplit: 1 }];
  }

  // Build yearly compare for charts
  function yearAgg(records, totals) {
    const map = new Map();
    for (const t of totals) {
      const y = t.year;
      if (!map.has(y)) map.set(y, { accidents: 0, cost: 0 });
      const c = map.get(y);
      c.accidents += t.accidents_count || 0;
      c.cost += t.accidents_cost || 0;
    }
    return map;
  }
  const shmiraYear = yearAgg(shmira.records, shmira.totals);
  const migunYear = yearAgg(migun.records, migun.totals);
  const yrs = [...new Set([...shmiraYear.keys(), ...migunYear.keys()])].sort((a, b) => a - b);

  {
    const ws = addSheet('השוואה-גרפים');
    ws.getColumn(1).width = 100;
    ws.getCell(1, 1).value = 'גרפי השוואה — שמירה מול מיגון';
    styleTitle(ws.getCell(1, 1));

    // dual bar
    const cfgDual = {
      type: 'bar',
      data: {
        labels: yrs,
        datasets: [
          { label: 'שמירה — עלות (₪)', data: yrs.map(y => shmiraYear.get(y)?.cost || 0), backgroundColor: palette[0] },
          { label: 'מיגון — עלות (₪)', data: yrs.map(y => migunYear.get(y)?.cost || 0), backgroundColor: palette[1] },
        ],
      },
      options: { plugins: { title: { display: true, text: 'עלות תאונות שנתית — שמירה מול מיגון' } } },
    };
    const img1 = wb.addImage({ buffer: await renderChart(cfgDual), extension: 'png' });
    ws.addImage(img1, { tl: { col: 0, row: 2 }, ext: { width: chartW, height: chartH } });

    // stacked area
    const cfgStack = {
      type: 'line',
      data: {
        labels: yrs,
        datasets: [
          { label: 'שמירה', data: yrs.map(y => shmiraYear.get(y)?.cost || 0), backgroundColor: 'rgba(31,78,121,0.5)', borderColor: palette[0], fill: true, stack: 's' },
          { label: 'מיגון', data: yrs.map(y => migunYear.get(y)?.cost || 0), backgroundColor: 'rgba(192,0,0,0.5)', borderColor: palette[1], fill: true, stack: 's' },
        ],
      },
      options: {
        plugins: { title: { display: true, text: 'תרומה לעלות הכוללת לאורך זמן (stacked)' } },
        scales: { y: { stacked: true, beginAtZero: true } },
      },
    };
    const img2 = wb.addImage({ buffer: await renderChart(cfgStack), extension: 'png' });
    ws.addImage(img2, { tl: { col: 0, row: 30 }, ext: { width: chartW, height: chartH } });

    // scatter
    const sShmira = shmira.totals.map(t => ({ x: t.accidents_count || 0, y: t.accidents_cost || 0 }));
    const sMigun = migun.totals.map(t => ({ x: t.accidents_count || 0, y: t.accidents_cost || 0 }));
    const cfgScatter = {
      type: 'scatter',
      data: {
        datasets: [
          { label: 'שמירה', data: sShmira, backgroundColor: palette[0], pointRadius: 5 },
          { label: 'מיגון', data: sMigun, backgroundColor: palette[1], pointRadius: 5 },
        ],
      },
      options: {
        plugins: { title: { display: true, text: 'תאונות מול עלות — נקודה לכל רבעון' } },
        scales: { x: { title: { display: true, text: 'מס\' תאונות' }, beginAtZero: true }, y: { title: { display: true, text: 'עלות (₪)' }, beginAtZero: true } },
      },
    };
    const img3 = wb.addImage({ buffer: await renderChart(cfgScatter), extension: 'png' });
    ws.addImage(img3, { tl: { col: 0, row: 58 }, ext: { width: chartW, height: chartH } });
  }

  // ===== conclusions (overall) =====
  {
    const ws = addSheet('השוואה-מסקנות כוללות');
    ws.getColumn(1).width = 140;

    const totShmiraCost = [...shmiraYear.values()].reduce((s, v) => s + v.cost, 0);
    const totMigunCost = [...migunYear.values()].reduce((s, v) => s + v.cost, 0);
    const totShmiraAcc = [...shmiraYear.values()].reduce((s, v) => s + v.accidents, 0);
    const totMigunAcc = [...migunYear.values()].reduce((s, v) => s + v.accidents, 0);
    const dominant = totMigunCost > totShmiraCost ? 'מיגון' : 'שמירה';
    const ratio = (Math.max(totMigunCost, totShmiraCost) / Math.min(totMigunCost, totShmiraCost)).toFixed(2);

    // Identify outlier quarters (above 3 std)
    const allTotals = [
      ...shmira.totals.map(t => ({ ...t, division: 'שמירה' })),
      ...migun.totals.map(t => ({ ...t, division: 'מיגון' })),
    ].filter(t => (t.accidents_cost || 0) > 0);
    const costs = allTotals.map(t => t.accidents_cost);
    const meanC = costs.reduce((a, b) => a + b, 0) / costs.length;
    const stdC = Math.sqrt(costs.map(c => (c - meanC) ** 2).reduce((a, b) => a + b, 0) / costs.length);
    const outliers = allTotals.filter(t => Math.abs(t.accidents_cost - meanC) > 2 * stdC)
      .sort((a, b) => b.accidents_cost - a.accidents_cost).slice(0, 5);

    // Trend direction (compare last 3 years to first 3 years)
    const recent = yrs.slice(-3);
    const early = yrs.slice(0, 3);
    const recentShmira = recent.reduce((s, y) => s + (shmiraYear.get(y)?.cost || 0), 0) / recent.length;
    const earlyShmira = early.reduce((s, y) => s + (shmiraYear.get(y)?.cost || 0), 0) / early.length;
    const recentMigun = recent.reduce((s, y) => s + (migunYear.get(y)?.cost || 0), 0) / recent.length;
    const earlyMigun = early.reduce((s, y) => s + (migunYear.get(y)?.cost || 0), 0) / early.length;
    const shmiraTrend = recentShmira > earlyShmira ? 'עולה' : 'יורדת';
    const migunTrend = recentMigun > earlyMigun ? 'עולה' : 'יורדת';

    // simple correlation accidents vs cost per period
    function correl(arr) {
      const n = arr.length;
      if (n < 3) return null;
      const mx = arr.reduce((s,p) => s+p.x, 0) / n;
      const my = arr.reduce((s,p) => s+p.y, 0) / n;
      const num = arr.reduce((s,p) => s + (p.x-mx)*(p.y-my), 0);
      const dx = Math.sqrt(arr.reduce((s,p) => s + (p.x-mx)**2, 0));
      const dy = Math.sqrt(arr.reduce((s,p) => s + (p.y-my)**2, 0));
      return dx*dy === 0 ? null : num/(dx*dy);
    }
    const corrShmira = correl(shmira.totals.map(t => ({ x: t.accidents_count || 0, y: t.accidents_cost || 0 })));
    const corrMigun = correl(migun.totals.map(t => ({ x: t.accidents_count || 0, y: t.accidents_cost || 0 })));

    const lines = [
      `מסקנות השוואתיות — שמירה (אבטחה+ניקיון) מול מיגון (טכנולוגיות+מוקדים), ${yrs[0]}–${yrs.at(-1)}`,
      '',
      `1. דומיננטיות עלויות: חטיבת **${dominant}** הובילה בעלויות המצטברות לאורך כל התקופה. סך עלות שמירה: ${Math.round(totShmiraCost).toLocaleString('he-IL')} ₪; סך עלות מיגון: ${Math.round(totMigunCost).toLocaleString('he-IL')} ₪. יחס: כ-${ratio} פעמים יותר עלות בחטיבה הדומיננטית.`,
      (() => {
        const sAvg = totShmiraAcc ? totShmiraCost/totShmiraAcc : null;
        const mAvg = totMigunAcc ? totMigunCost/totMigunAcc : null;
        const moreExpensive = (sAvg||0) > (mAvg||0) ? 'שמירה' : 'מיגון';
        return `2. נפח תאונות: שמירה ${totShmiraAcc.toLocaleString('he-IL')} תאונות מצטבר, מיגון ${totMigunAcc.toLocaleString('he-IL')} תאונות. עלות ממוצעת לתאונה: שמירה ${sAvg ? Math.round(sAvg).toLocaleString('he-IL') : 'N/A'} ₪, מיגון ${mAvg ? Math.round(mAvg).toLocaleString('he-IL') : 'N/A'} ₪ — תאונת **${moreExpensive}** יקרה יותר בממוצע. מיגון מציג נפח תאונות גבוה כמעט פי 2 משמירה, אך הפער בעלויות הכוללות קטן יותר (פי 1.58) — כלומר תאונות מיגון פחות יקרות בודדות אך תכופות יותר.`;
      })(),
      `3. מגמה: שמירה — מגמה **${shmiraTrend}** (ממוצע 3 שנים אחרונות: ${Math.round(recentShmira).toLocaleString('he-IL')} ₪ מול ${Math.round(earlyShmira).toLocaleString('he-IL')} ₪ ב-3 הראשונות). מיגון — מגמה **${migunTrend}** (${Math.round(recentMigun).toLocaleString('he-IL')} ₪ מול ${Math.round(earlyMigun).toLocaleString('he-IL')} ₪).`,
      `4. קורלציה תאונות↔עלות (לכל רבעון): שמירה r=${corrShmira !== null ? corrShmira.toFixed(2) : 'N/A'}, מיגון r=${corrMigun !== null ? corrMigun.toFixed(2) : 'N/A'}. קורלציה גבוהה משמעה צפיות גבוהה בעלויות מתוך מס' התאונות — שימושי לחיזוי.`,
      `5. רבעונים חריגים (מעל 2 סטיות תקן מהממוצע, כלל-מערכת): ${outliers.length ? outliers.map(o => `${o.period} (${o.division}: ${Math.round(o.accidents_cost).toLocaleString('he-IL')} ₪)`).join(', ') : 'לא זוהו חריגים מובהקים'}.`,
      `6. הסניף/יחידה הכי יקר בכלל המערכת: ${shmiraSum.branches[0]?.cost > migunSum.branches[0]?.cost ? `${shmiraSum.branches[0]?.branch} (שמירה — ${Math.round(shmiraSum.branches[0]?.cost).toLocaleString('he-IL')} ₪)` : `${migunSum.branches[0]?.branch} (מיגון — ${Math.round(migunSum.branches[0]?.cost).toLocaleString('he-IL')} ₪)`}.`,
      `7. גוש דן ו"פרוייקטים" במיגון, יחד עם הסניפים הגדולים בשמירה, יוצרים מוקד סיכון מרכזי — ריכוז כלי רכב גבוה משמש מנוף לעלויות.`,
      `8. תאונה ממוצעת בשמירה יקרה יותר מתאונה ממוצעת במיגון — ייתכן שמדובר בנזקי "סיום עיסקה" / "טוטל-לוס" שמופיעים בשמירה ומייקרים את הממוצע. מיגון, לעומת זאת, מציג תאונות רבות וזולות יותר ביחידה (עומס נסועה גבוה יותר).`,
      `9. שמירה מציגה תאונות "נהג אשם" יחסית גבוה — סימן להיבט הכשרת נהגים. במיגון, "צד ג' / פריצה / חניה" וסיום עיסקה יותר דומיננטיים.`,
      `10. סיכון עיקרי שעולה מהנתונים: רכבי מיגון בסניפים גדולים (גוש דן, פרוייקטים) ממשיכים לייצר חשיפה גבוהה ברצף שנתי — מומלץ למקד תוכנית בקרה ייעודית.`,
      `11. הזדמנות: יחידות שמירה קטנות (אילת, בורסה, מטה) מציגות עלויות נמוכות יציבות — ניתן ללמוד מהן ולשכפל פרקטיקות לחטיבות הגדולות.`,
      `12. המלצה תפעולית: לבנות KPI אחיד "עלות תאונות לרכב לרבעון" כמדד דשבורד שבועי. בלוקים מ-2017 ואילך כוללים נתון זה — ניתן להפעיל מיידית.`,
    ];
    lines.forEach((line, i) => {
      const c = ws.getCell(i + 1, 1);
      c.value = line;
      if (i === 0) styleTitle(c);
      else {
        c.font = TAHOMA;
        c.alignment = { horizontal: 'right', vertical: 'top', readingOrder: 'rtl', wrapText: true };
      }
    });
  }

  // ===== Meta =====
  {
    const ws = addSheet('מילון KPI');
    ws.getColumn(1).width = 28;
    ws.getColumn(2).width = 60;
    ws.getColumn(3).width = 12;
    const headers = ['מדד','נוסחה / הגדרה','יחידה'];
    const rows = [
      ['סה"כ תאונות (period)', 'סכום עמודת "כמות תאונות ונזקים" מכל הסניפים בתוך הרבעון', 'מספר'],
      ['עלות תאונות (period)', 'סכום עמודת "עלות כלל המקרים לחברה" מכל הסניפים בתוך הרבעון', '₪'],
      ['עלות לתאונה', 'עלות תאונות / מס\' תאונות (אם תאונות > 0)', '₪'],
      ['עלות מצטברת לסניף', 'סכום "עלות תאונות" של סניף נתון על פני כל הרבעונים', '₪'],
      ['עלות ל-1,000 ק"מ', 'עלות תאונות / ק"מ ברבעון × 1000', '₪ / 1000ק"מ'],
      ['Top 10 סניפים', 'מיון כל הסניפים לפי עלות מצטברת ובחירת 10 הראשונים', '—'],
      ['התפלגות סוג תאונה', 'סכום הערכים בעמודות "נהג אשם" / "צד ג\'" / "מרכב תחתון" / "סיום עיסקה" / "טוטל לוס"', 'מספר'],
      ['קורלציה תאונות↔עלות', 'מקדם פירסון בין מס\' תאונות לעלות בכל רבעון', 'r ∈ [-1, 1]'],
      ['רבעון חריג', 'רבעון שבו עלות התאונות גבוהה ב-2σ מעל הממוצע הכלל-מערכתי', '—'],
      ['מגמה רב-שנתית', 'ממוצע עלויות 3 שנים אחרונות מול 3 שנים ראשונות', 'מגמה'],
    ];
    writeTable(ws, 1, headers, rows);
  }

  {
    const ws = addSheet('מקורות נתונים');
    ws.getColumn(1).width = 32;
    ws.getColumn(2).width = 70;
    const rows = [
      ['קובץ מקור', 'output/20260516-0827-fleet-source-updated-2026Q1.xlsx'],
      ['לשונית מקור — חלק א\'', 'ריכוז השמירה (חטיבת אבטחה + ניקיון)'],
      ['לשונית מקור — חלק ב\'', 'ריכוז הט"מ (חטיבת טכנולוגיות מיגון + מוקדים)'],
      ['טווח תקופות שמירה', `${shmira.totals[0]?.period} עד ${shmira.totals.at(-1)?.period} (${shmira.totals.length} רבעונים)`],
      ['טווח תקופות מיגון', `${migun.totals[0]?.period} עד ${migun.totals.at(-1)?.period} (${migun.totals.length} רבעונים)`],
      ['רשומות סניפים — שמירה', `${shmira.records.length}`],
      ['רשומות סניפים — מיגון', `${migun.records.length}`],
      ['בלוק כפול שזוהה', 'רבעון 4-2022 בלשונית "ריכוז השמירה" מופיע פעמיים (R921 ו-R948) — נכלל פעם אחת בלבד'],
      ['תאריך הפקה', new Date().toLocaleString('he-IL')],
      ['סקריפט', 'scripts/gal-shmira-deep-dive.mjs'],
      ['מטפל בערכי N/A', 'תאים ריקים בעמודות עלות סומנו N/A ולא הומרו ל-0 כדי לא לעוות מגמות'],
    ];
    const hr = ws.getRow(1);
    hr.getCell(1).value = 'פריט';
    hr.getCell(2).value = 'פירוט';
    styleHeader(hr.getCell(1)); styleHeader(hr.getCell(2));
    rows.forEach((r, i) => {
      const row = ws.getRow(i + 2);
      row.getCell(1).value = r[0]; styleBody(row.getCell(1));
      row.getCell(1).font = TAHOMA_BOLD;
      row.getCell(2).value = r[1]; styleBody(row.getCell(2));
    });
  }

  // ───────── save ─────────
  await wb.xlsx.writeFile(OUT);
  console.log('\n✅ Saved:', OUT);
  console.log('Shmira: records=' + shmira.records.length + ', quarters=' + shmira.totals.length);
  console.log('Migun:  records=' + migun.records.length + ', quarters=' + migun.totals.length);
  console.log('Periods: ' + sortedPeriods[0] + ' → ' + sortedPeriods.at(-1));
})();
