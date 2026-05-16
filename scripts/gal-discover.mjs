import * as XLSX from 'xlsx';
import fs from 'node:fs';

const file = 'input/תאונות  רבעון 1-2026.xls';
const buf = fs.readFileSync(file);
const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, codepage: 1255 });

console.log('=== SHEETS ===');
console.log('Count:', wb.SheetNames.length);
console.log('Names:', wb.SheetNames);

for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
  console.log(`\n=== SHEET: ${name} ===`);
  console.log('Rows:', rows.length);
  if (rows.length > 0) {
    console.log('Columns:', Object.keys(rows[0]));
    console.log('First 5 rows:');
    console.log(JSON.stringify(rows.slice(0, 5), null, 2));
  }
}
