#!/usr/bin/env node
// Gal — Q1 2026 focused analysis + Year-over-Year Q1 2025 vs Q1 2026.
// Reads the updated source workbook (output/20260516-0827-fleet-source-updated-2026Q1.xlsx),
// parses both summary sheets, extracts only Q1 2025 and Q1 2026 blocks per division,
// and produces a single Excel workbook with executive summary, division focus,
// division comparison, YoY per division, unified YoY, and "significant movers" sheets.

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
const OUT = path.join(ROOT, 'output', `${stamp}-fleet-q1-yoy-2025-vs-2026.xlsx`);

// ───────── style constants ─────────
const TAHOMA = { name: 'Tahoma', size: 11 };
const TAHOMA_BOLD = { name: 'Tahoma', size: 11, bold: true };
const TAHOMA_TITLE = { name: 'Tahoma', size: 14, bold: true };
const TAHOMA_KPI_LABEL = { name: 'Tahoma', size: 10, color: { argb: 'FF555555' } };
const TAHOMA_KPI_VALUE = { name: 'Tahoma', size: 18, bold: true, color: { argb: 'FF1F4E79' } };
const fmtMoney = '#,##0\\ ₪';
const fmtInt = '#,##0';
const fmtPct = '0.0%';
const fmtMoneyDec = '#,##0.00\\ ₪';
const palette = ['#1F4E79','#C00000','#548235','#BF8F00','#7030A0','#2E75B6','#A52A2A','#385723','#806000','#4472C4'];

// ───────── helpers ─────────
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
  const s = String(label).replace(/\s+/g, ' ').trim();
  let m = s.match(/רבעון\s*([1-4])\s*שנת\s*(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4])\s*[-–]\s*(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4])\s+(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*[-–]\s*(\d{4})\s*[-–]\s*([1-4])/);
  if (m) return { q: +m[2], y: +m[1] };
  m = s.match(/רבעון\s*[-–]\s*([1-4])\s*[-–]\s*(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  m = s.match(/רבעון\s*([1-4]).*?(\d{4})/);
  if (m) return { q: +m[1], y: +m[2] };
  return null;
}

// Parse a single sheet → array of period blocks, each with records[] and total{}
function parseSheet(ws) {
  const rows = ws.rowCount;
  const headers = [];
  for (let r = 1; r <= rows; r++) {
    const v = cellVal(ws.getRow(r).getCell(2));
    if (isQuarterHeader(v)) {
      const p = parsePeriodLabel(v);
      if (p) headers.push({ row: r, label: String(v).trim(), ...p });
    }
  }
  const blocks = [];
  const seen = new Set();
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    const periodKey = `${h.y}-Q${h.q}`;
    const nextRow = i + 1 < headers.length ? headers[i + 1].row : rows + 1;
    let headerRow = null;
    for (let r = h.row + 1; r <= Math.min(h.row + 5, rows); r++) {
      const v = String(cellVal(ws.getRow(r).getCell(2))).trim();
      if (v === 'סניף' || v === '.' || /סניף/.test(v)) { headerRow = r; break; }
    }
    if (!headerRow) continue;
    if (seen.has(periodKey)) continue;
    seen.add(periodKey);
    const records = [];
    let total = null;
    for (let r = headerRow + 1; r < nextRow; r++) {
      const branch = String(cellVal(ws.getRow(r).getCell(2))).trim();
      if (!branch) continue;
      if (branch.includes('סה"כ') || branch.includes('סה״כ') || branch.includes("סה'כ")) {
        const row = ws.getRow(r);
        total = {
          period: periodKey, year: h.y, quarter: h.q,
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
        };
        break;
      }
      const row = ws.getRow(r);
      records.push({
        period: periodKey, year: h.y, quarter: h.q, branch,
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
      });
    }
    blocks.push({ period: periodKey, year: h.y, quarter: h.q, label: h.label, headerRow: h.row, records, total });
  }
  return blocks;
}

// ───────── styling helpers ─────────
function styleHeader(cell, color = 'FF1F4E79') {
  cell.font = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
  cell.alignment = { horizontal: 'right', vertical: 'middle', readingOrder: 'rtl', wrapText: true };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
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
function writeTable(ws, startRow, headers, rows, opts = {}) {
  const { moneyCols = [], pctCols = [], intCols = [], moneyDecCols = [], colorByDelta = [], headerColor = 'FF1F4E79' } = opts;
  const hr = ws.getRow(startRow);
  headers.forEach((h, i) => {
    const c = hr.getCell(i + 1);
    c.value = h;
    styleHeader(c, headerColor);
  });
  hr.height = 30;
  rows.forEach((r, ri) => {
    const row = ws.getRow(startRow + 1 + ri);
    r.forEach((v, ci) => {
      const c = row.getCell(ci + 1);
      c.value = (v === null || v === undefined) ? 'נתון חסר' : v;
      styleBody(c);
      if (moneyCols.includes(ci)) c.numFmt = fmtMoney;
      else if (moneyDecCols.includes(ci)) c.numFmt = fmtMoneyDec;
      else if (pctCols.includes(ci)) c.numFmt = fmtPct;
      else if (intCols.includes(ci)) c.numFmt = fmtInt;
      if (v === null || v === undefined) {
        c.font = { name: 'Tahoma', size: 11, italic: true, color: { argb: 'FF888888' } };
      }
      // Delta colouring
      if (colorByDelta.includes(ci) && typeof v === 'number') {
        if (v > 0) c.font = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FFC00000' } };
        else if (v < 0) c.font = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FF548235' } };
      }
    });
  });
  return startRow + 1 + rows.length;
}

// ───────── KPI computation for a Q1 block ─────────
function blockStats(records, total) {
  // Use total row if it exists, else sum records
  const acc = total?.accidents_count ?? records.reduce((s,r)=>s+(r.accidents_count||0),0);
  const cost = total?.accidents_cost ?? records.reduce((s,r)=>s+(r.accidents_cost||0),0);
  const vehicles = total?.vehicles ?? records.reduce((s,r)=>s+(r.vehicles||0),0);
  const km = total?.km ?? records.reduce((s,r)=>s+(r.km||0),0);
  const branches = records.filter(r => r.branch && !/^\.$/.test(r.branch)).length;
  // per-branch cost array for stats
  const costs = records.map(r => r.accidents_cost || 0).filter(v => v > 0);
  const mean = costs.length ? costs.reduce((a,b)=>a+b,0)/costs.length : 0;
  const variance = costs.length ? costs.reduce((s,v)=>s+(v-mean)**2,0)/costs.length : 0;
  const stdev = Math.sqrt(variance);
  const sorted = [...costs].sort((a,b)=>a-b);
  const median = sorted.length ? (sorted.length % 2 ? sorted[(sorted.length-1)/2] : (sorted[sorted.length/2-1]+sorted[sorted.length/2])/2) : 0;
  const avg_per_accident = acc ? cost/acc : null;

  // type mix
  const types = {
    'נהג החברה אשם': records.reduce((s,r)=>s+(r.driver_at_fault||0),0),
    'צד ג\' / פריצה / חניה': records.reduce((s,r)=>s+(r.third_party||0),0),
    'נזק מרכב תחתון': records.reduce((s,r)=>s+(r.undercarriage||0),0),
    'נזקי סיום עיסקה': records.reduce((s,r)=>s+(r.end_of_deal||0),0),
    'טוטל-לוס': records.reduce((s,r)=>s+(r.total_loss||0),0),
  };
  // top 5 branches by cost
  const top5 = [...records]
    .filter(r => r.branch)
    .sort((a,b)=>(b.accidents_cost||0)-(a.accidents_cost||0))
    .slice(0,5)
    .map(r => ({ branch: r.branch, accidents: r.accidents_count||0, cost: r.accidents_cost||0, vehicles: r.vehicles||0 }));

  // outliers = branches with cost > mean + 3*std
  const threshold = mean + 3*stdev;
  const outliers = records
    .filter(r => (r.accidents_cost||0) > threshold && (r.accidents_cost||0) > 0)
    .map(r => ({ branch: r.branch, cost: r.accidents_cost, accidents: r.accidents_count }));

  return { acc, cost, vehicles, km, branches, mean, median, stdev, avg_per_accident, types, top5, outliers, threshold };
}

// ───────── YoY computation ─────────
function buildYoY(rec2025, rec2026) {
  const map2025 = new Map();
  for (const r of rec2025) {
    if (!r.branch || /^\.$/.test(r.branch)) continue;
    map2025.set(r.branch.trim(), r);
  }
  const map2026 = new Map();
  for (const r of rec2026) {
    if (!r.branch || /^\.$/.test(r.branch)) continue;
    map2026.set(r.branch.trim(), r);
  }
  const allBranches = new Set([...map2025.keys(), ...map2026.keys()]);
  const rows = [];
  for (const b of allBranches) {
    const a = map2025.get(b);
    const c = map2026.get(b);
    const acc25 = a?.accidents_count;
    const acc26 = c?.accidents_count;
    const cost25 = a?.accidents_cost;
    const cost26 = c?.accidents_cost;
    let dAcc = null, dAccPct = null, dCost = null, dCostPct = null, note = '';
    if (acc25 != null && acc26 != null) {
      dAcc = acc26 - acc25;
      if (acc25 === 0 && acc26 > 0) { dAccPct = null; note = `+${acc26} תאונות חדשות`; }
      else if (acc25 > 0) dAccPct = (acc26 - acc25) / acc25;
      else if (acc25 === 0 && acc26 === 0) dAccPct = 0;
    } else if (acc25 == null) note = 'לא הופיע 2025';
    else if (acc26 == null) note = 'לא הופיע 2026';

    if (cost25 != null && cost26 != null) {
      dCost = cost26 - cost25;
      if (cost25 === 0 && cost26 > 0) dCostPct = null;
      else if (cost25 > 0) dCostPct = (cost26 - cost25) / cost25;
      else if (cost25 === 0 && cost26 === 0) dCostPct = 0;
    }
    rows.push({ branch: b, acc25, acc26, dAcc, dAccPct, cost25, cost26, dCost, dCostPct, note });
  }
  return rows.sort((x,y) => (y.cost26||0)-(x.cost26||0));
}

// ───────── chart rendering ─────────
const chartW = 900, chartH = 500;
const chartCanvas = new ChartJSNodeCanvas({
  width: chartW, height: chartH, backgroundColour: 'white',
  chartCallback: (ChartJS) => {
    ChartJS.defaults.font.family = 'Tahoma';
    ChartJS.defaults.font.size = 12;
  },
});
async function renderChart(config) { return await chartCanvas.renderToBuffer(config, 'image/png'); }

// ───────── main ─────────
(async () => {
  console.log('Reading source workbook…');
  const src = new ExcelJS.Workbook();
  await src.xlsx.readFile(SRC);
  const wsS = src.getWorksheet('ריכוז השמירה');
  const wsM = src.getWorksheet('ריכוז הט"מ');

  const shmiraBlocks = parseSheet(wsS);
  const migunBlocks = parseSheet(wsM);
  console.log(`Shmira blocks: ${shmiraBlocks.length}, Migun blocks: ${migunBlocks.length}`);

  // Find Q1 2025 + Q1 2026 in each division
  const findBlock = (blocks, y, q) => blocks.find(b => b.year === y && b.quarter === q);
  const sh2026 = findBlock(shmiraBlocks, 2026, 1);
  const sh2025 = findBlock(shmiraBlocks, 2025, 1);
  const mg2026 = findBlock(migunBlocks, 2026, 1);
  const mg2025 = findBlock(migunBlocks, 2025, 1);

  if (!sh2026 || !sh2025 || !mg2026 || !mg2025) {
    console.error('Missing required blocks. Found:', {
      sh2026: !!sh2026, sh2025: !!sh2025, mg2026: !!mg2026, mg2025: !!mg2025,
    });
    process.exit(1);
  }
  console.log('Found all four Q1 blocks. Computing stats…');

  const shStats26 = blockStats(sh2026.records, sh2026.total);
  const shStats25 = blockStats(sh2025.records, sh2025.total);
  const mgStats26 = blockStats(mg2026.records, mg2026.total);
  const mgStats25 = blockStats(mg2025.records, mg2025.total);

  const shYoY = buildYoY(sh2025.records, sh2026.records);
  const mgYoY = buildYoY(mg2025.records, mg2026.records);

  // Unified group totals
  const grp25 = {
    acc: shStats25.acc + mgStats25.acc,
    cost: shStats25.cost + mgStats25.cost,
    vehicles: shStats25.vehicles + mgStats25.vehicles,
    km: shStats25.km + mgStats25.km,
  };
  const grp26 = {
    acc: shStats26.acc + mgStats26.acc,
    cost: shStats26.cost + mgStats26.cost,
    vehicles: shStats26.vehicles + mgStats26.vehicles,
    km: shStats26.km + mgStats26.km,
  };
  const grpDelta = {
    acc: grp26.acc - grp25.acc,
    accPct: grp25.acc ? (grp26.acc - grp25.acc)/grp25.acc : null,
    cost: grp26.cost - grp25.cost,
    costPct: grp25.cost ? (grp26.cost - grp25.cost)/grp25.cost : null,
  };

  // significant movers (only branches with comparable data)
  const allYoY = [
    ...shYoY.map(r => ({ ...r, division: 'שמירה' })),
    ...mgYoY.map(r => ({ ...r, division: 'מיגון' })),
  ];
  const improved = allYoY.filter(r => r.dAccPct != null && r.dAccPct <= -0.20);
  const worsened = allYoY.filter(r => r.dAccPct != null && r.dAccPct >= 0.20);

  console.log('Building workbook…');
  const wb = new ExcelJS.Workbook();
  wb.creator = 'גל — Q1 YoY 2025↔2026';
  wb.created = new Date();
  function addSheet(name) {
    const ws = wb.addWorksheet(name, { properties: { defaultRowHeight: 18 } });
    setRTL(ws);
    return ws;
  }

  // ============================================================
  // Sheet 1: A-סיכום מנהלים
  // ============================================================
  {
    const ws = addSheet('A-סיכום מנהלים');
    for (let i = 1; i <= 8; i++) ws.getColumn(i).width = 18;
    ws.getColumn(1).width = 24;

    ws.getCell(1, 1).value = 'דוח Q1 2026 + השוואת YoY מול Q1 2025';
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 6);

    ws.getCell(2, 1).value = 'הופק על-ידי גל (Excel Fleet Analyst) — קבוצת G1';
    ws.getCell(2, 1).font = TAHOMA_KPI_LABEL;
    ws.getCell(2, 1).alignment = { horizontal: 'right', readingOrder: 'rtl' };
    ws.mergeCells(2, 1, 2, 6);

    // ── KPI Cards row (Q1 2026 group)
    let r = 4;
    ws.getCell(r, 1).value = 'KPI ראשיים — Q1 2026 קבוצה'; styleTitle(ws.getCell(r, 1));
    r += 2;
    const cards = [
      ['סה"כ תאונות Q1 2026', grp26.acc, fmtInt],
      ['סה"כ עלות Q1 2026', grp26.cost, fmtMoney],
      ['רכבים פעילים', grp26.vehicles, fmtInt],
      ['ק"מ נסיעה ברבעון', grp26.km, fmtInt],
      ['עלות ממוצעת לתאונה', grp26.acc ? grp26.cost/grp26.acc : null, fmtMoney],
      ['עלות לרכב', grp26.vehicles ? grp26.cost/grp26.vehicles : null, fmtMoney],
    ];
    cards.forEach((card, i) => {
      const col = (i % 3) * 2 + 1;
      const row = r + Math.floor(i/3) * 3;
      const lbl = ws.getCell(row, col);
      lbl.value = card[0]; lbl.font = TAHOMA_KPI_LABEL;
      lbl.alignment = { horizontal: 'right', readingOrder: 'rtl' };
      ws.mergeCells(row, col, row, col+1);
      const val = ws.getCell(row+1, col);
      val.value = card[1] ?? 'נתון חסר';
      val.font = TAHOMA_KPI_VALUE;
      val.numFmt = card[2];
      val.alignment = { horizontal: 'right', readingOrder: 'rtl' };
      ws.mergeCells(row+1, col, row+1, col+1);
    });

    // YoY summary table
    r += 8;
    ws.getCell(r, 1).value = 'השוואת YoY (Q1 2025 → Q1 2026) — קבוצה'; styleTitle(ws.getCell(r, 1));
    r += 2;
    r = writeTable(ws, r,
      ['חטיבה','Q1-2025 תאונות','Q1-2026 תאונות','Δ תאונות','Δ%','Q1-2025 עלות','Q1-2026 עלות','Δ עלות','Δ%'],
      [
        ['שמירה', shStats25.acc, shStats26.acc, shStats26.acc - shStats25.acc, shStats25.acc ? (shStats26.acc - shStats25.acc)/shStats25.acc : null,
          shStats25.cost, shStats26.cost, shStats26.cost - shStats25.cost, shStats25.cost ? (shStats26.cost - shStats25.cost)/shStats25.cost : null],
        ['מיגון', mgStats25.acc, mgStats26.acc, mgStats26.acc - mgStats25.acc, mgStats25.acc ? (mgStats26.acc - mgStats25.acc)/mgStats25.acc : null,
          mgStats25.cost, mgStats26.cost, mgStats26.cost - mgStats25.cost, mgStats25.cost ? (mgStats26.cost - mgStats25.cost)/mgStats25.cost : null],
        ['סה"כ קבוצה', grp25.acc, grp26.acc, grpDelta.acc, grpDelta.accPct, grp25.cost, grp26.cost, grpDelta.cost, grpDelta.costPct],
      ],
      { intCols: [1, 2, 3], moneyCols: [5, 6, 7], pctCols: [4, 8], colorByDelta: [3, 4, 7, 8] });

    // 3 key findings
    r += 2;
    ws.getCell(r, 1).value = '3 ממצאי-על'; styleTitle(ws.getCell(r, 1));
    r += 1;
    const verdictAcc = grpDelta.accPct == null ? 'ללא נתון' : (grpDelta.accPct < 0 ? 'הצי משתפר ✓' : 'הצי מחמיר ✗');
    const verdictCost = grpDelta.costPct == null ? 'ללא נתון' : (grpDelta.costPct < 0 ? 'עלויות יורדות ✓' : 'עלויות עולות ✗');
    const findings = [
      `1. כיוון כללי: ${verdictAcc}. נפח התאונות השתנה ב-${grpDelta.accPct != null ? (grpDelta.accPct*100).toFixed(1)+'%' : 'לא ידוע'} (${grp25.acc} → ${grp26.acc} תאונות).`,
      `2. עלויות: ${verdictCost}. סך עלות התאונות השתנה ב-${grpDelta.costPct != null ? (grpDelta.costPct*100).toFixed(1)+'%' : 'לא ידוע'} (${Math.round(grp25.cost).toLocaleString('he-IL')} ₪ → ${Math.round(grp26.cost).toLocaleString('he-IL')} ₪).`,
      `3. ${improved.length} סניפים השתפרו >20% YoY, ${worsened.length} סניפים החמירו >20% YoY. המסר העיקרי: ${verdictAcc === 'הצי משתפר ✓' && verdictCost === 'עלויות יורדות ✓' ? 'השיפור אחיד ומובהק.' : verdictAcc === 'הצי מחמיר ✗' && verdictCost === 'עלויות עולות ✗' ? 'מגמת הרעה רוחבית — דורש התערבות.' : 'תמונה מעורבת — נדרשת התבוננות פר חטיבה.'}`,
    ];
    findings.forEach(f => {
      r++;
      const c = ws.getCell(r, 1);
      c.value = f;
      c.font = TAHOMA;
      c.alignment = { horizontal: 'right', vertical: 'top', readingOrder: 'rtl', wrapText: true };
      ws.mergeCells(r, 1, r, 9);
      ws.getRow(r).height = 32;
    });

    // Key chart: YoY group comparison
    r += 2;
    ws.getCell(r, 1).value = 'גרף מפתח — Q1 2025 vs Q1 2026 (תאונות + עלות, פר חטיבה)';
    styleTitle(ws.getCell(r, 1));
    r += 1;
    const cfg = {
      type: 'bar',
      data: {
        labels: ['שמירה — תאונות','מיגון — תאונות','קבוצה — תאונות'],
        datasets: [
          { label: 'Q1 2025', data: [shStats25.acc, mgStats25.acc, grp25.acc], backgroundColor: palette[5] },
          { label: 'Q1 2026', data: [shStats26.acc, mgStats26.acc, grp26.acc], backgroundColor: palette[1] },
        ],
      },
      options: { plugins: { title: { display: true, text: 'מס\' תאונות Q1 — 2025 מול 2026' } } },
    };
    const img = wb.addImage({ buffer: await renderChart(cfg), extension: 'png' });
    ws.addImage(img, { tl: { col: 0, row: r }, ext: { width: chartW, height: chartH } });
  }

  // ============================================================
  // Sheet 2: A-Q1-2026-שמירה
  // ============================================================
  function buildDivisionFocusSheet(name, stats, block, color) {
    const ws = addSheet(name);
    for (let i = 1; i <= 6; i++) ws.getColumn(i).width = 20;
    ws.getColumn(1).width = 28;

    ws.getCell(1, 1).value = `${name} — נתוני בסיס ופילוחים`;
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 6);

    let r = 3;
    // Base KPIs
    r = writeTable(ws, r,
      ['מדד','ערך'],
      [
        ['סה"כ תאונות', stats.acc],
        ['סה"כ עלות (₪)', stats.cost],
        ['רכבים פעילים', stats.vehicles],
        ['ק"מ נסיעה', stats.km],
        ['סניפים פעילים', stats.branches],
        ['עלות ממוצעת לתאונה', stats.avg_per_accident],
        ['עלות לרכב', stats.vehicles ? stats.cost/stats.vehicles : null],
        ['ממוצע עלות לסניף', stats.mean],
        ['חציון עלות לסניף', stats.median],
        ['סטיית תקן (עלות לסניף)', stats.stdev],
        ['סף חריג סטטיסטי (mean+3σ)', stats.threshold],
      ], { headerColor: color, intCols: [], moneyCols: [] });
    // numFmt: rows with non-money intvals
    // (we set general; OK)

    r += 2;
    ws.getCell(r, 1).value = '5 הסניפים היקרים ביותר (לפי עלות תאונות)';
    styleTitle(ws.getCell(r, 1));
    r += 1;
    r = writeTable(ws, r,
      ['סניף','תאונות','עלות תאונות (₪)','רכבים','עלות לרכב (₪)'],
      stats.top5.map(t => [t.branch, t.accidents, t.cost, t.vehicles, t.vehicles ? t.cost/t.vehicles : null]),
      { intCols: [1, 3], moneyCols: [2, 4], headerColor: color });

    r += 2;
    ws.getCell(r, 1).value = 'חלוקה לפי סוג תאונה';
    styleTitle(ws.getCell(r, 1));
    r += 1;
    const typeTotal = Object.values(stats.types).reduce((a,b)=>a+b,0);
    r = writeTable(ws, r,
      ['סוג תאונה','כמות','שיעור'],
      Object.entries(stats.types).map(([t, c]) => [t, c, typeTotal ? c/typeTotal : 0]),
      { intCols: [1], pctCols: [2], headerColor: color });

    r += 2;
    ws.getCell(r, 1).value = `חריגים סטטיסטיים (סניפים עם עלות > מקסימום ${Math.round(stats.threshold).toLocaleString('he-IL')} ₪)`;
    styleTitle(ws.getCell(r, 1));
    r += 1;
    if (stats.outliers.length === 0) {
      ws.getCell(r, 1).value = 'לא זוהו סניפים חריגים מעל הסף.';
      ws.getCell(r, 1).font = TAHOMA;
      ws.getCell(r, 1).alignment = { horizontal: 'right', readingOrder: 'rtl' };
    } else {
      r = writeTable(ws, r, ['סניף','תאונות','עלות (₪)'],
        stats.outliers.map(o => [o.branch, o.accidents, o.cost]),
        { intCols: [1], moneyCols: [2], headerColor: color });
    }
    return ws;
  }
  buildDivisionFocusSheet('A-Q1-2026-שמירה', shStats26, sh2026, 'FF1F4E79');
  buildDivisionFocusSheet('A-Q1-2026-מיגון', mgStats26, mg2026, 'FFC00000');

  // ============================================================
  // Sheet 4: A-Q1-2026-השוואה (Shmira vs Migun)
  // ============================================================
  {
    const ws = addSheet('A-Q1-2026-השוואה');
    for (let i = 1; i <= 5; i++) ws.getColumn(i).width = 22;

    ws.getCell(1, 1).value = 'השוואת חטיבות Q1 2026 — שמירה מול מיגון';
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 5);

    let r = 3;
    r = writeTable(ws, r,
      ['מדד','שמירה','מיגון','הפרש (שמירה-מיגון)','% של שמירה מתוך הקבוצה'],
      [
        ['תאונות', shStats26.acc, mgStats26.acc, shStats26.acc - mgStats26.acc, grp26.acc ? shStats26.acc/grp26.acc : 0],
        ['עלות (₪)', shStats26.cost, mgStats26.cost, shStats26.cost - mgStats26.cost, grp26.cost ? shStats26.cost/grp26.cost : 0],
        ['רכבים', shStats26.vehicles, mgStats26.vehicles, shStats26.vehicles - mgStats26.vehicles, grp26.vehicles ? shStats26.vehicles/grp26.vehicles : 0],
        ['ק"מ', shStats26.km, mgStats26.km, shStats26.km - mgStats26.km, grp26.km ? shStats26.km/grp26.km : 0],
        ['סניפים', shStats26.branches, mgStats26.branches, shStats26.branches - mgStats26.branches, null],
        ['עלות ממוצעת לתאונה (₪)', shStats26.avg_per_accident, mgStats26.avg_per_accident,
          (shStats26.avg_per_accident||0) - (mgStats26.avg_per_accident||0), null],
        ['עלות לרכב (₪)', shStats26.vehicles ? shStats26.cost/shStats26.vehicles : null,
          mgStats26.vehicles ? mgStats26.cost/mgStats26.vehicles : null, null, null],
      ],
      { moneyCols: [], pctCols: [4] });

    // Type-mix side-by-side
    r += 2;
    ws.getCell(r, 1).value = 'תמהיל סוגי תאונות — Q1 2026'; styleTitle(ws.getCell(r, 1));
    r += 1;
    const types = Object.keys(shStats26.types);
    const totalSh = Object.values(shStats26.types).reduce((a,b)=>a+b,0);
    const totalMg = Object.values(mgStats26.types).reduce((a,b)=>a+b,0);
    r = writeTable(ws, r,
      ['סוג תאונה','שמירה — כמות','שמירה — %','מיגון — כמות','מיגון — %'],
      types.map(t => [t, shStats26.types[t], totalSh ? shStats26.types[t]/totalSh : 0, mgStats26.types[t], totalMg ? mgStats26.types[t]/totalMg : 0]),
      { intCols: [1, 3], pctCols: [2, 4] });

    // Chart
    r += 2;
    const cfg = {
      type: 'bar',
      data: {
        labels: ['תאונות','עלות (₪/1000)'],
        datasets: [
          { label: 'שמירה', data: [shStats26.acc, shStats26.cost/1000], backgroundColor: palette[0] },
          { label: 'מיגון', data: [mgStats26.acc, mgStats26.cost/1000], backgroundColor: palette[1] },
        ],
      },
      options: {
        plugins: { title: { display: true, text: 'Q1 2026 — שמירה מול מיגון (עלות מוצגת באלפי ₪)' } },
      },
    };
    const img = wb.addImage({ buffer: await renderChart(cfg), extension: 'png' });
    ws.addImage(img, { tl: { col: 0, row: r }, ext: { width: chartW, height: chartH } });
  }

  // ============================================================
  // Sheets 5+6: B-YoY-שמירה, B-YoY-מיגון
  // ============================================================
  async function buildYoYDivisionSheet(name, rows, color, divName) {
    const ws = addSheet(name);
    for (let i = 1; i <= 10; i++) ws.getColumn(i).width = 16;
    ws.getColumn(1).width = 24;
    ws.getColumn(10).width = 22;

    ws.getCell(1, 1).value = `${divName}: השוואת Q1 2025 → Q1 2026 פר סניף`;
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 10);

    let r = 3;
    r = writeTable(ws, r,
      ['סניף','Q1-2025 תאונות','Q1-2026 תאונות','Δ תאונות','Δ% תאונות','Q1-2025 עלות (₪)','Q1-2026 עלות (₪)','Δ עלות (₪)','Δ% עלות','הערה'],
      rows.map(r2 => [
        r2.branch,
        r2.acc25, r2.acc26, r2.dAcc, r2.dAccPct,
        r2.cost25, r2.cost26, r2.dCost, r2.dCostPct,
        r2.note,
      ]),
      { intCols: [1, 2, 3], pctCols: [4, 8], moneyCols: [5, 6, 7],
        colorByDelta: [3, 4, 7, 8], headerColor: color });

    // Twin-bars chart — branch accidents 2025 vs 2026 (top 12 by 2026 cost)
    const top = rows.filter(r2 => r2.acc25 != null || r2.acc26 != null).slice(0, 12);
    r += 2;
    const cfg = {
      type: 'bar',
      data: {
        labels: top.map(t => t.branch),
        datasets: [
          { label: 'Q1 2025', data: top.map(t => t.acc25 ?? 0), backgroundColor: palette[5] },
          { label: 'Q1 2026', data: top.map(t => t.acc26 ?? 0), backgroundColor: color === 'FF1F4E79' ? palette[0] : palette[1] },
        ],
      },
      options: { plugins: { title: { display: true, text: `${divName}: תאונות פר סניף — Q1 2025 vs Q1 2026` } } },
    };
    const img = wb.addImage({ buffer: await renderChart(cfg), extension: 'png' });
    ws.addImage(img, { tl: { col: 0, row: r }, ext: { width: chartW, height: chartH } });
  }
  await buildYoYDivisionSheet('B-YoY-שמירה', shYoY, 'FF1F4E79', 'שמירה');
  await buildYoYDivisionSheet('B-YoY-מיגון', mgYoY, 'FFC00000', 'מיגון');

  // ============================================================
  // Sheet 7: C-YoY-מאוחד
  // ============================================================
  {
    const ws = addSheet('C-YoY-מאוחד');
    for (let i = 1; i <= 6; i++) ws.getColumn(i).width = 20;
    ws.getColumn(1).width = 26;

    ws.getCell(1, 1).value = 'C — YoY מאוחד (Q1 2025 → Q1 2026)';
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 6);

    let r = 3;
    r = writeTable(ws, r,
      ['רמה','Q1-2025 תאונות','Q1-2026 תאונות','Δ%','Q1-2025 עלות','Q1-2026 עלות','Δ%'],
      [
        ['שמירה', shStats25.acc, shStats26.acc, shStats25.acc ? (shStats26.acc-shStats25.acc)/shStats25.acc : null,
          shStats25.cost, shStats26.cost, shStats25.cost ? (shStats26.cost-shStats25.cost)/shStats25.cost : null],
        ['מיגון', mgStats25.acc, mgStats26.acc, mgStats25.acc ? (mgStats26.acc-mgStats25.acc)/mgStats25.acc : null,
          mgStats25.cost, mgStats26.cost, mgStats25.cost ? (mgStats26.cost-mgStats25.cost)/mgStats25.cost : null],
        ['קבוצה (סה"כ)', grp25.acc, grp26.acc, grpDelta.accPct, grp25.cost, grp26.cost, grpDelta.costPct],
      ],
      { intCols: [1, 2], pctCols: [3, 6], moneyCols: [4, 5], colorByDelta: [3, 6] });

    // type-mix YoY group level
    r += 2;
    ws.getCell(r, 1).value = 'תמהיל סוגי תאונות — קבוצה (Q1 2025 vs Q1 2026)';
    styleTitle(ws.getCell(r, 1));
    r += 1;
    const types = Object.keys(shStats26.types);
    const grpTypes25 = Object.fromEntries(types.map(t => [t, (shStats25.types[t]||0)+(mgStats25.types[t]||0)]));
    const grpTypes26 = Object.fromEntries(types.map(t => [t, (shStats26.types[t]||0)+(mgStats26.types[t]||0)]));
    const total25 = Object.values(grpTypes25).reduce((a,b)=>a+b,0);
    const total26 = Object.values(grpTypes26).reduce((a,b)=>a+b,0);
    r = writeTable(ws, r,
      ['סוג תאונה','2025 — כמות','2025 — %','2026 — כמות','2026 — %','שינוי %'],
      types.map(t => {
        const v25 = grpTypes25[t]; const v26 = grpTypes26[t];
        const share25 = total25 ? v25/total25 : 0;
        const share26 = total26 ? v26/total26 : 0;
        const delta = share26 - share25;
        return [t, v25, share25, v26, share26, delta];
      }),
      { intCols: [1, 3], pctCols: [2, 4, 5], colorByDelta: [5] });

    // Chart: mixed bars (acc) + line (cost)
    r += 2;
    const cfg = {
      type: 'bar',
      data: {
        labels: ['Q1 2025','Q1 2026'],
        datasets: [
          { type: 'bar', label: 'שמירה — תאונות', data: [shStats25.acc, shStats26.acc], backgroundColor: palette[0], yAxisID: 'y' },
          { type: 'bar', label: 'מיגון — תאונות', data: [mgStats25.acc, mgStats26.acc], backgroundColor: palette[1], yAxisID: 'y' },
          { type: 'line', label: 'סה"כ עלות (₪)', data: [grp25.cost, grp26.cost], borderColor: palette[3], backgroundColor: palette[3], yAxisID: 'y1', tension: 0.2, pointRadius: 6, fill: false },
        ],
      },
      options: {
        plugins: { title: { display: true, text: 'YoY מאוחד — תאונות (עמודות) + עלות (קו)' } },
        scales: {
          y: { beginAtZero: true, title: { display: true, text: 'מס\' תאונות' } },
          y1: { beginAtZero: true, position: 'left', grid: { drawOnChartArea: false }, title: { display: true, text: 'עלות (₪)' } },
        },
      },
    };
    const img = wb.addImage({ buffer: await renderChart(cfg), extension: 'png' });
    ws.addImage(img, { tl: { col: 0, row: r }, ext: { width: chartW, height: chartH } });
  }

  // ============================================================
  // Sheets 8+9: C-משופרים מובהקים / C-החמירו מובהקים
  // ============================================================
  function buildMoversSheet(name, items, headerColor, title) {
    const ws = addSheet(name);
    for (let i = 1; i <= 8; i++) ws.getColumn(i).width = 17;
    ws.getColumn(1).width = 24;
    ws.getCell(1, 1).value = title;
    styleTitle(ws.getCell(1, 1));
    ws.mergeCells(1, 1, 1, 8);
    let r = 3;
    if (items.length === 0) {
      ws.getCell(r, 1).value = 'אין סניפים העונים על הקריטריון.';
      ws.getCell(r, 1).font = TAHOMA;
      ws.getCell(r, 1).alignment = { horizontal: 'right', readingOrder: 'rtl' };
      return;
    }
    r = writeTable(ws, r,
      ['סניף','חטיבה','Q1-2025 תאונות','Q1-2026 תאונות','Δ תאונות','Δ%','Q1-2025 עלות','Q1-2026 עלות'],
      items.map(x => [x.branch, x.division, x.acc25, x.acc26, x.dAcc, x.dAccPct, x.cost25, x.cost26]),
      { intCols: [2, 3, 4], pctCols: [5], moneyCols: [6, 7], colorByDelta: [4, 5], headerColor });
  }
  improved.sort((a,b)=>a.dAccPct - b.dAccPct);
  worsened.sort((a,b)=>b.dAccPct - a.dAccPct);
  buildMoversSheet('C-משופרים מובהקים', improved, 'FF548235', 'סניפים שהשתפרו ב-20% או יותר YoY (תאונות)');
  buildMoversSheet('C-החמירו מובהקים', worsened, 'FFC00000', 'סניפים שהחמירו ב-20% או יותר YoY (תאונות)');

  // ============================================================
  // Sheet 10: מילון KPI
  // ============================================================
  {
    const ws = addSheet('מילון KPI');
    ws.getColumn(1).width = 30;
    ws.getColumn(2).width = 70;
    ws.getColumn(3).width = 14;
    writeTable(ws, 1, ['מדד','נוסחה / הגדרה','יחידה'], [
      ['סה"כ תאונות (רבעון)', 'סכום העמודה "כמות תאונות ונזקים" מכל סניף ברבעון (נלקח משורת "סה"כ" כאשר קיימת)', 'מספר'],
      ['סה"כ עלות (רבעון)', 'סכום העמודה "עלות כלל המקרים לחברה" מכל סניף ברבעון', '₪'],
      ['עלות ממוצעת לתאונה', 'עלות תאונות / מס\' תאונות (אם תאונות > 0)', '₪'],
      ['עלות לרכב', 'עלות תאונות / מס\' רכבים פעילים', '₪'],
      ['ממוצע עלות לסניף', 'mean של עלות התאונות פר סניף (סניפים עם עלות > 0)', '₪'],
      ['חציון עלות לסניף', 'median של עלות התאונות פר סניף (סניפים עם עלות > 0)', '₪'],
      ['סטיית תקן לסניף', 'std של עלות התאונות פר סניף', '₪'],
      ['סף חריג סטטיסטי', 'mean + 3·σ — סניף עם עלות מעל הסף מוגדר חריג', '₪'],
      ['Δ תאונות', 'Q1-2026 תאונות – Q1-2025 תאונות', 'מספר'],
      ['Δ% תאונות', '(Q1-2026 – Q1-2025) / Q1-2025; אם Q1-2025=0 ויש 2026 → "X תאונות חדשות"', '%'],
      ['Δ% עלות', '(Q1-2026 עלות – Q1-2025 עלות) / Q1-2025 עלות', '%'],
      ['משופר מובהק', 'סניף שבו Δ% תאונות ≤ –20% (השתפר ביותר מ-20%)', '—'],
      ['החמיר מובהק', 'סניף שבו Δ% תאונות ≥ +20%', '—'],
      ['נתון חסר', 'סניף שמופיע ב-2025 בלבד או ב-2026 בלבד — לא ניתן לחשב Δ%', '—'],
      ['Σ ברמת קבוצה', 'סכימת נומרטור ודנומינטור משתי החטיבות; לא ממוצע של ממוצעים', '—'],
    ], { intCols: [], moneyCols: [], pctCols: [] });
  }

  // ============================================================
  // Sheet 11: מקורות נתונים
  // ============================================================
  {
    const ws = addSheet('מקורות נתונים');
    ws.getColumn(1).width = 34;
    ws.getColumn(2).width = 80;
    const rows = [
      ['קובץ מקור', 'output/20260516-0827-fleet-source-updated-2026Q1.xlsx (עותק מעודכן של תאונות 1-2026)'],
      ['לשונית — שמירה', `ריכוז השמירה — בלוקים: ${shmiraBlocks.length} רבעונים`],
      ['לשונית — מיגון', `ריכוז הט"מ — בלוקים: ${migunBlocks.length} רבעונים`],
      ['Q1 2025 שמירה — שורת כותרת', `שורה ${sh2025.headerRow} ("${sh2025.label}")`],
      ['Q1 2025 מיגון — שורת כותרת', `שורה ${mg2025.headerRow} ("${mg2025.label}")`],
      ['Q1 2026 שמירה — שורת כותרת', `שורה ${sh2026.headerRow} ("${sh2026.label}")`],
      ['Q1 2026 מיגון — שורת כותרת', `שורה ${mg2026.headerRow} ("${mg2026.label}")`],
      ['שמירה — סניפים Q1 2025', `${sh2025.records.length} רשומות`],
      ['שמירה — סניפים Q1 2026', `${sh2026.records.length} רשומות`],
      ['מיגון — סניפים Q1 2025', `${mg2025.records.length} רשומות`],
      ['מיגון — סניפים Q1 2026', `${mg2026.records.length} רשומות`],
      ['תאריך הפקה', new Date().toLocaleString('he-IL')],
      ['סקריפט', 'scripts/gal-q1-yoy-2025-2026.mjs'],
      ['טיפול בנתון חסר', 'תא ריק בעלות סומן "נתון חסר"; Δ% מ-0 לערך חיובי מוצג כ"X תאונות חדשות"'],
      ['הערה — נתונים אגרגטיביים', 'הניתוח מבוסס רק על לשוניות הריכוז (שורות מסכמות פר סניף); לא נטענה לשונית "תאונות 1-2026" הגלמית.'],
    ];
    const hr = ws.getRow(1);
    hr.getCell(1).value = 'פריט'; hr.getCell(2).value = 'פירוט';
    styleHeader(hr.getCell(1)); styleHeader(hr.getCell(2));
    rows.forEach((r, i) => {
      const row = ws.getRow(i + 2);
      row.getCell(1).value = r[0]; styleBody(row.getCell(1)); row.getCell(1).font = TAHOMA_BOLD;
      row.getCell(2).value = r[1]; styleBody(row.getCell(2));
    });
  }

  // ───────── save ─────────
  await wb.xlsx.writeFile(OUT);
  console.log('\n✅ Saved:', OUT);

  // ───────── print report payload ─────────
  const fmt = n => Math.round(n).toLocaleString('he-IL');
  const pct = p => p == null ? 'N/A' : (p*100).toFixed(1) + '%';
  console.log('\n────── REPORT PAYLOAD ──────');
  console.log(`OUT_FILE=${path.relative(ROOT, OUT)}`);
  console.log(`SH_2026 acc=${shStats26.acc} cost=${fmt(shStats26.cost)} branches=${shStats26.branches}`);
  console.log(`SH_2025 acc=${shStats25.acc} cost=${fmt(shStats25.cost)} branches=${shStats25.branches}`);
  console.log(`MG_2026 acc=${mgStats26.acc} cost=${fmt(mgStats26.cost)} branches=${mgStats26.branches}`);
  console.log(`MG_2025 acc=${mgStats25.acc} cost=${fmt(mgStats25.cost)} branches=${mgStats25.branches}`);
  console.log(`GRP_2026 acc=${grp26.acc} cost=${fmt(grp26.cost)}`);
  console.log(`GRP_2025 acc=${grp25.acc} cost=${fmt(grp25.cost)}`);
  console.log(`GRP DELTA acc=${grpDelta.acc} (${pct(grpDelta.accPct)}) cost=${fmt(grpDelta.cost)} (${pct(grpDelta.costPct)})`);
  console.log(`SHMIRA YoY: dAcc%=${pct(shStats25.acc ? (shStats26.acc-shStats25.acc)/shStats25.acc : null)}  dCost%=${pct(shStats25.cost ? (shStats26.cost-shStats25.cost)/shStats25.cost : null)}`);
  console.log(`MIGUN  YoY: dAcc%=${pct(mgStats25.acc ? (mgStats26.acc-mgStats25.acc)/mgStats25.acc : null)}  dCost%=${pct(mgStats25.cost ? (mgStats26.cost-mgStats25.cost)/mgStats25.cost : null)}`);
  console.log(`IMPROVED >20%: ${improved.length}  WORSENED >20%: ${worsened.length}`);
  console.log('Top improved (top 5):');
  improved.slice(0, 5).forEach(i => console.log(`  ${i.division} | ${i.branch}: ${i.acc25}→${i.acc26} (${pct(i.dAccPct)})`));
  console.log('Top worsened (top 5):');
  worsened.slice(0, 5).forEach(i => console.log(`  ${i.division} | ${i.branch}: ${i.acc25}→${i.acc26} (${pct(i.dAccPct)})`));
  console.log('SHMIRA top 5 Q1 2026:');
  shStats26.top5.forEach(t => console.log(`  ${t.branch}: ${t.accidents} acc, ${fmt(t.cost)} ₪`));
  console.log('MIGUN top 5 Q1 2026:');
  mgStats26.top5.forEach(t => console.log(`  ${t.branch}: ${t.accidents} acc, ${fmt(t.cost)} ₪`));
  console.log('SHMIRA outliers (mean+3σ):', shStats26.outliers.map(o => `${o.branch}(${fmt(o.cost)})`).join(', ') || 'אין');
  console.log('MIGUN  outliers (mean+3σ):', mgStats26.outliers.map(o => `${o.branch}(${fmt(o.cost)})`).join(', ') || 'אין');
  // type mix shift
  const types = Object.keys(shStats26.types);
  for (const t of types) {
    const v25 = (shStats25.types[t]||0)+(mgStats25.types[t]||0);
    const v26 = (shStats26.types[t]||0)+(mgStats26.types[t]||0);
    console.log(`TYPE ${t}: 2025=${v25} → 2026=${v26}`);
  }
})();
