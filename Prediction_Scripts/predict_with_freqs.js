// predict_with_freqs_exceljs_sep_year.js
// Usage:
//   node Prediction_Scripts\predict_with_freqs.js Test_Data\weekly.xlsx Test_Data\monthly.xlsx
// Outputs:
//   ./<scriptBase>/weekly_year_predictions.xlsx
//   ./<scriptBase>/monthly_year_predictions.xlsx

const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const fs = require('fs');
const path = require('path');

if (process.argv.length < 4) {
  console.error('Usage: node predict_with_freqs_exceljs_sep_year.js weekly.xlsx monthly.xlsx');
  process.exit(1);
}

const weeklyFile = process.argv[2];
const monthlyFile = process.argv[3];

// ---------- create output folder based on script name ----------
const scriptBase = path.basename(process.argv[1], path.extname(process.argv[1]));
const outputFolder = path.join(process.cwd(), scriptBase);
if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder, { recursive: true });
console.log('Output folder:', outputFolder);

// ---------- helpers ----------
const is3Digits = s => {
  if (s === null || s === undefined) return false;
  const str = String(s).trim();
  const m = str.match(/(\d{1,3})$/);
  if (!m) return false;
  return /^[0-9]{3}$/.test(m[1].padStart(3,'0'));
};
const normalize3 = s => {
  const str = String(s||'').trim();
  const m = str.match(/(\d{1,3})$/);
  if (!m) return null;
  return m[1].padStart(3,'0');
};
function mkCounter() { return { pos0: {}, pos1: {}, pos2: {}, triplets: {}, total: 0 }; }
function addToCounter(counter, triplet, times = 1) {
  counter.total = (counter.total || 0) + times;
  counter.pos0[triplet[0]] = (counter.pos0[triplet[0]] || 0) + times;
  counter.pos1[triplet[1]] = (counter.pos1[triplet[1]] || 0) + times;
  counter.pos2[triplet[2]] = (counter.pos2[triplet[2]] || 0) + times;
  counter.triplets[triplet] = (counter.triplets[triplet] || 0) + times;
}
function allTriplets() { const out = []; for (let i=0;i<1000;i++) out.push(String(i).padStart(3,'0')); return out; }
const TRIPLETS = allTriplets();

// ---------- read sheet into matrix ----------
async function sheetToMatrix(pathArg, preferredNames = []) {
  if (!fs.existsSync(pathArg)) throw new Error('File not found: ' + pathArg);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(pathArg);
  let sheet = null;
  for (const n of preferredNames) {
    const s = wb.getWorksheet(n);
    if (s) { sheet = s; break; }
  }
  if (!sheet) sheet = wb.worksheets[0];
  const matrix = [];
  sheet.eachRow((row) => {
    const vals = row.values ? row.values.slice(1).map(v => (v === undefined || v === null) ? '' : String(v)) : [];
    matrix.push(vals);
  });
  return matrix;
}

// ---------- parse weekly matrix ----------
function parseWeeklyMatrix(mat) {
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const dayNames = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
  let headerIdx = -1;
  for (let i=0;i<Math.min(mat.length,6);i++){
    const row = mat[i].map(c => up(c));
    let matches = 0;
    for (const dn of dayNames) if (row.some(cell => cell.includes(dn))) matches++;
    if (matches >= 3) { headerIdx = i; break; }
  }
  let start = 0; let colMap = null;
  if (headerIdx >= 0) {
    start = headerIdx + 1;
    colMap = {};
    const header = mat[headerIdx].map(c => up(c));
    header.forEach((cell, idx) => {
      if (!cell) return;
      if (cell.includes('SUN')) colMap[idx] = 0;
      else if (cell.includes('MON')) colMap[idx] = 1;
      else if (cell.includes('TUE')) colMap[idx] = 2;
      else if (cell.includes('WED')) colMap[idx] = 3;
      else if (cell.includes('THU')) colMap[idx] = 4;
      else if (cell.includes('FRI')) colMap[idx] = 5;
      else if (cell.includes('SAT')) colMap[idx] = 6;
    });
  } else {
    const maxCols = Math.max(...mat.map(r=>r.length));
    if (maxCols >= 7) {
      colMap = {}; for (let c=0;c<7;c++) colMap[c] = c; start = 0;
    } else {
      return [];
    }
  }

  const results = [];
  for (let r=start; r<mat.length; r++){
    const row = mat[r];
    if (!row) continue;
    for (const [col, weekday] of Object.entries(colMap)) {
      const cidx = Number(col);
      const cell = row[cidx];
      if (!cell) continue;
      if (is3Digits(cell)) results.push({ triplet: normalize3(cell), weekday: Number(weekday) });
    }
  }
  return results;
}

// ---------- parse monthly matrix ----------
function parseMonthlyMatrix(mat) {
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const headerIdx = mat.findIndex(row => row && row.some(c => up(c).includes('DATE')));
  const headerRow = mat[ headerIdx >= 0 ? headerIdx : 0 ].map(c => up(c));
  const months = { 'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12 };
  const colToMonth = {};
  headerRow.forEach((cell, idx) => {
    if (!cell) return;
    for (const key of Object.keys(months)) if (cell.includes(key)) { colToMonth[idx] = months[key]; break; }
    const short = cell.slice(0,3);
    if (!colToMonth[idx] && months[short]) colToMonth[idx] = months[short];
  });
  let dateCol = headerRow.findIndex(c => up(c).includes('DATE'));
  if (dateCol < 0) dateCol = 0;
  const start = (headerIdx >= 0) ? headerIdx + 1 : 1;
  const results = [];
  for (let r=start; r<mat.length; r++) {
    const row = mat[r];
    if (!row) continue;
    const dateVal = row[dateCol];
    const dayNum = parseInt(String(dateVal||'').trim(), 10);
    if (!dayNum || dayNum < 1 || dayNum > 31) continue;
    for (const [cStr, mon] of Object.entries(colToMonth)) {
      const cidx = Number(cStr);
      const cell = row[cidx];
      if (!cell) continue;
      if (is3Digits(cell)) results.push({ triplet: normalize3(cell), month: Number(mon), day: dayNum });
    }
  }
  return results;
}

// ---------- build counters separate for weekly & monthly ----------
async function ingestWeeklyOnly(weeklyPath) {
  const matrix = await sheetToMatrix(weeklyPath, ['weekly','Week','Sheet1','Sheet']);
  const parsed = parseWeeklyMatrix(matrix);
  const weekCounters = {}; // 0..6
  const overall = mkCounter();
  for (const p of parsed) {
    if (!weekCounters[p.weekday]) weekCounters[p.weekday] = mkCounter();
    addToCounter(weekCounters[p.weekday], p.triplet);
    addToCounter(overall, p.triplet);
  }
  return { weekCounters, overall };
}

async function ingestMonthlyOnly(monthlyPath) {
  const matrix = await sheetToMatrix(monthlyPath, ['monthly','Month','Sheet1','Sheet']);
  const parsed = parseMonthlyMatrix(matrix);
  const monthCounters = {}; // 1..12
  const monthDayTrip = {}; // map 'm-d' -> {triplet: count}
  const overall = mkCounter();
  for (const p of parsed) {
    if (!monthCounters[p.month]) monthCounters[p.month] = mkCounter();
    addToCounter(monthCounters[p.month], p.triplet);
    const key = `${p.month}-${p.day}`;
    if (!monthDayTrip[key]) monthDayTrip[key] = {};
    monthDayTrip[key][p.triplet] = (monthDayTrip[key][p.triplet]||0) + 1;
    addToCounter(overall, p.triplet);
  }
  // convert monthDayTrip to counters
  const monthDayCounters = {};
  for (const [key, map] of Object.entries(monthDayTrip)) {
    const cnt = mkCounter();
    for (const [trip, c] of Object.entries(map)) {
      cnt.triplets[trip] = c;
      cnt.total += c;
      // positional counts
      for (let i=0;i<c;i++) {
        cnt.pos0[trip[0]] = (cnt.pos0[trip[0]]||0) + 1;
        cnt.pos1[trip[1]] = (cnt.pos1[trip[1]]||0) + 1;
        cnt.pos2[trip[2]] = (cnt.pos2[trip[2]]||0) + 1;
      }
    }
    monthDayCounters[key] = cnt;
  }
  return { monthCounters, monthDayCounters, overall };
}

// ---------- build positional/triplet rows ----------
function buildPositionalRows(counter) {
  const total = counter.total || 0;
  const rows = [['Digit','Count_pos0','Prob_pos0','Count_pos1','Prob_pos1','Count_pos2','Prob_pos2']];
  for (let d=0; d<=9; d++) {
    const k = String(d);
    const c0 = counter.pos0[k]||0, c1 = counter.pos1[k]||0, c2 = counter.pos2[k]||0;
    rows.push([k, c0, total? c0/total : 0, c1, total? c1/total : 0, c2, total? c2/total : 0]);
  }
  rows.push(['TOTAL', total]);
  return rows;
}
function buildTripletRows(counter) {
  const total = counter.total || 0;
  const arr = Object.entries(counter.triplets || {}).map(([t,cnt]) => ({ t, cnt, prob: total? cnt/total : 0 }));
  arr.sort((a,b) => b.cnt - a.cnt);
  const rows = [['Triplet','Count','Probability']];
  for (const e of arr) rows.push([e.t, e.cnt, e.prob]);
  rows.push(['TOTAL_OBS', total]);
  return rows;
}

// ---------- ranking helper used for predictions ----------
function buildProbMaps(counter) {
  const total = counter.total || 1;
  const p0 = {}, p1 = {}, p2 = {};
  for (let d=0; d<=9; d++) {
    const k = String(d);
    p0[k] = (counter.pos0[k] || 0) / total;
    p1[k] = (counter.pos1[k] || 0) / total;
    p2[k] = (counter.pos2[k] || 0) / total;
  }
  return { p0, p1, p2, triplets: counter.triplets || {}, total: counter.total || 0 };
}

function rankCandidatesFromCounter(counter, topKPerPos=6, topN=50) {
  const combos = new Map();
  for (const [t,c] of Object.entries(counter.triplets || {})) {
    const score = (counter.total ? c / counter.total : 0);
    combos.set(t, Math.max(combos.get(t) || 0, score));
  }
  const pm = buildProbMaps(counter);
  const sortDigits = pmap => Object.keys(pmap).sort((a,b) => pmap[b] - pmap[a]);
  const top0 = sortDigits(pm.p0).slice(0, topKPerPos);
  const top1 = sortDigits(pm.p1).slice(0, topKPerPos);
  const top2 = sortDigits(pm.p2).slice(0, topKPerPos);
  for (const a of top0) for (const b of top1) for (const c of top2) {
    const n = `${a}${b}${c}`;
    const score = (pm.p0[a]||0)*(pm.p1[b]||0)*(pm.p2[c]||0);
    combos.set(n, Math.max(combos.get(n) || 0, score));
  }
  const arr = Array.from(combos.entries()).map(([num,score]) => ({ num, score }));
  arr.sort((x,y) => y.score - x.score);
  return arr.slice(0, topN).map(x => ({ num: x.num, score: x.score, count: (counter.triplets && counter.triplets[x.num]) || 0 }));
}

// ---------- generate next 365 dates (starting tomorrow) ----------
function nextNDates(n=365) {
  const out = [];
  let day = dayjs().startOf('day').add(1, 'day');
  for (let i=0;i<n;i++) { out.push(day.format('YYYY-MM-DD')); day = day.add(1, 'day'); }
  return out;
}

// ---------- produce weekly workbook (weekly-only analysis) ----------
async function produceWeeklyWorkbook(weeklyPath, outFileName) {
  const { weekCounters, overall } = await ingestWeeklyOnly(weeklyPath);

  const wb = new ExcelJS.Workbook();

  // positional_overall
  const posAll = buildPositionalRows(overall);
  const wsPosAll = wb.addWorksheet('positional_overall');
  posAll.forEach(r => wsPosAll.addRow(r));

  // positional_weekday blocks
  const wsPosWeek = wb.addWorksheet('positional_weekday');
  for (let w=0; w<=6; w++) {
    wsPosWeek.addRow([`WEEKDAY_${w}`]);
    const block = buildPositionalRows(weekCounters[w] || mkCounter());
    block.forEach(r => wsPosWeek.addRow(r));
    wsPosWeek.addRow([]);
  }

  // triplets_weekday
  const wsTripWeek = wb.addWorksheet('triplets_weekday');
  wsTripWeek.addRow(['Weekday','Triplet','Count','Probability']);
  for (let w=0; w<=6; w++) {
    const cnt = weekCounters[w] || mkCounter();
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c]) => ({t,c,p: tot? c/tot : 0})).sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripWeek.addRow([w, it.t, it.c, it.p]);
  }

  // predictions sheet for next 365 days (weekly-only)
  const wsPred = wb.addWorksheet('predictions');
  wsPred.addRow(['Date','Weekday','ObservationsUsed','TopCandidates']);
  const dates = nextNDates(365);
  for (const d of dates) {
    const wd = dayjs(d).day();
    const cnt = weekCounters[wd] || mkCounter();
    const candidates = rankCandidatesFromCounter(cnt, 6, 50);
    const obs = cnt.total || 0;
    const topText = candidates.map(c => `${c.num} (score:${(c.score*100).toFixed(4)}%, count:${c.count})`).join(' | ');
    wsPred.addRow([d, wd, obs, topText]);
  }

  const outPath = path.join(outputFolder, outFileName);
  await wb.xlsx.writeFile(outPath);
  console.log('Weekly workbook written:', outPath);
}

// ---------- produce monthly workbook (monthly-only analysis) ----------
async function produceMonthlyWorkbook(monthlyPath, outFileName) {
  const { monthCounters, monthDayCounters, overall } = await ingestMonthlyOnly(monthlyPath);

  const wb = new ExcelJS.Workbook();

  // positional_month blocks
  const wsPosMonth = wb.addWorksheet('positional_month');
  for (let m=1; m<=12; m++) {
    wsPosMonth.addRow([`MONTH_${m}`]);
    const block = buildPositionalRows(monthCounters[m] || mkCounter());
    block.forEach(r => wsPosMonth.addRow(r));
    wsPosMonth.addRow([]);
  }

  // triplets_month
  const wsTripMonth = wb.addWorksheet('triplets_month');
  wsTripMonth.addRow(['Month','Triplet','Count','Probability']);
  for (let m=1; m<=12; m++) {
    const cnt = monthCounters[m] || mkCounter();
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c]) => ({t,c,p: tot? c/tot : 0})).sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripMonth.addRow([m, it.t, it.c, it.p]);
  }

  // triplets_month_day
  const wsTripMonthDay = wb.addWorksheet('triplets_month_day');
  wsTripMonthDay.addRow(['Month','Day','Triplet','Count','Probability']);
  for (const key of Object.keys(monthDayCounters)) {
    const [m,d] = key.split('-').map(Number);
    const cnt = monthDayCounters[key];
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c]) => ({t,c,p: tot? c/tot : 0})).sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripMonthDay.addRow([m,d,it.t,it.c,it.p]);
  }

  // predictions sheet for next 365 days (monthly-only)
  const wsPred = wb.addWorksheet('predictions');
  wsPred.addRow(['Date','Month','Day','ObservationsUsed','SourceUsed','TopCandidates']);
  const dates = nextNDates(365);
  for (const d of dates) {
    const dt = dayjs(d);
    const m = dt.month() + 1;
    const dayNum = dt.date();
    const key = `${m}-${dayNum}`;
    let source = 'none', cnt = null;
    if (monthDayCounters[key] && (monthDayCounters[key].total || 0) > 0) {
      cnt = monthDayCounters[key]; source = 'month-day';
    } else if (monthCounters[m] && (monthCounters[m].total || 0) > 0) {
      cnt = monthCounters[m]; source = 'month';
    } else {
      cnt = overall; source = 'overall';
    }
    const candidates = rankCandidatesFromCounter(cnt, 6, 50);
    const obs = cnt.total || 0;
    const topText = candidates.map(c => `${c.num} (score:${(c.score*100).toFixed(4)}%, count:${c.count})`).join(' | ');
    wsPred.addRow([d, m, dayNum, obs, source, topText]);
  }

  const outPath = path.join(outputFolder, outFileName);
  await wb.xlsx.writeFile(outPath);
  console.log('Monthly workbook written:', outPath);
}

// ---------- run ----------
(async () => {
  try {
    await produceWeeklyWorkbook(weeklyFile, 'weekly_year_predictions.xlsx');
  } catch (err) {
    console.error('Weekly generation failed:', err.message);
  }
  try {
    await produceMonthlyWorkbook(monthlyFile, 'monthly_year_predictions.xlsx');
  } catch (err) {
    console.error('Monthly generation failed:', err.message);
  }
})();
