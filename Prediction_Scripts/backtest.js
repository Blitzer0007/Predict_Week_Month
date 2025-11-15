// backtest.js
// Node.js backtest for weekly & monthly lottery 3-digit predictions
// Usage: node Prediction_Scripts\backtest.js Test_Data\weekly.xlsx Test_Data\monthly.xlsx 
//
// - Reads weekly/monthly Excel (flexible formats).
// - Ignores 'XXX' (case-insensitive) and non-3-digit cells.
// - Performs expanding-window backtest (train on all previous records).
// - Computes Top-K hit rates, MRR, Brier score.
// - Writes per-case CSVs inside a folder named after this JS filename.

const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const fs = require('fs');
const path = require('path');

// ---------------- CONFIG ----------------
const MIN_TRAIN = 50;           // minimum training observations before starting evaluation
const ALPHA_TRIPLET = 1.0;      // Laplace smoothing for triplets (pseudo-count)
const ALPHA_POS = 1.0;          // Laplace smoothing for positional counts (per-digit)
const MIX_TRIPLET = 0.5;        // weight for triplet distribution vs positional-product (0..1)
const TOP_KS = [1,5,10,20];     // which top-K to compute hit rates for
const MAX_CLASSES = 1000;       // 000..999

// ---------------- create output folder based on script name ----------------
const scriptName = path.basename(process.argv[1], path.extname(process.argv[1]));
const outputFolder = path.join(process.cwd(), scriptName);
if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder, { recursive: true });
console.log('Outputs will be written to folder:', outputFolder);

// ---------------- helpers ----------------
const zeroPad3 = n => String(n).padStart(3,'0');
function allTriplets() { const out=[]; for (let i=0;i<1000;i++) out.push(zeroPad3(i)); return out; }
const TRIPLETS = allTriplets();

function isXXXorInvalid(s) {
  if (s === null || s === undefined) return true;
  const t = String(s).trim();
  if (t === '') return true;
  if (/^xxx$/i.test(t)) return true;
  const m = t.match(/(\d{1,3})$/);
  if (!m) return true;
  return false;
}
function extract3(s) {
  if (s === null || s === undefined) return null;
  const t = String(s).trim();
  const m = t.match(/(\d{1,3})$/);
  if (!m) return null;
  return m[1].padStart(3,'0');
}
function mkCounter() {
  return { total:0, pos0:Object.create(null), pos1:Object.create(null), pos2:Object.create(null), triplets:Object.create(null) };
}
function addToCounter(counter, trip, times=1) {
  counter.total += times;
  counter.pos0[trip[0]] = (counter.pos0[trip[0]]||0) + times;
  counter.pos1[trip[1]] = (counter.pos1[trip[1]]||0) + times;
  counter.pos2[trip[2]] = (counter.pos2[trip[2]]||0) + times;
  counter.triplets[trip] = (counter.triplets[trip]||0) + times;
}

// ---------------- Excel reading utils (flexible) ----------------
async function sheetToMatrix(filePath, preferredNames=[]) {
  if (!fs.existsSync(filePath)) throw new Error('File missing: ' + filePath);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  let sheet = null;
  for (const n of preferredNames) {
    const s = wb.getWorksheet(n);
    if (s) { sheet = s; break; }
  }
  if (!sheet) sheet = wb.worksheets[0];
  const mat = [];
  sheet.eachRow((row) => {
    const vals = row.values ? row.values.slice(1).map(v => (v===undefined||v===null)?'': String(v)) : [];
    mat.push(vals);
  });
  return mat;
}

// try to parse a flattened sheet containing Date & Number columns
function tryParseFlattened(mat) {
  if (!mat || mat.length < 2) return null;
  const headerRow = mat[0].map(c => String(c||'').toLowerCase());
  const dateCol = headerRow.findIndex(h => h.includes('date'));
  const numCol = headerRow.findIndex(h => /num|number|value|draw|result/.test(h));
  if (dateCol >= 0 && numCol >= 0) {
    const rows = [];
    for (let r=1; r<mat.length; r++) {
      const row = mat[r];
      const rawDate = row[dateCol];
      const rawNum = row[numCol];
      if (isXXXorInvalid(rawNum)) continue;
      const num = extract3(rawNum);
      if (!num) continue;
      let date = null;
      try {
        date = dayjs(String(rawDate));
        if (!date.isValid()) date = null;
      } catch(e) { date = null; }
      rows.push({ date, num, rawDate: rawDate });
    }
    return rows;
  }
  return null;
}

// parse weekly grid: tries to detect SUN..SAT header; looks for optional Week/Date column
function parseWeeklyGrid(mat) {
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const names = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
  let headerIdx = -1;
  for (let i=0;i<Math.min(6,mat.length); i++) {
    const row = mat[i].map(c => up(c));
    const matches = names.reduce((acc,nm) => acc + (row.some(cell => cell.includes(nm)) ? 1 : 0), 0);
    if (matches >= 3) { headerIdx = i; break; }
  }
  if (headerIdx < 0) {
    const maxCols = Math.max(...mat.map(r=>r.length));
    if (maxCols < 7) return [];
  }
  let startRow = (headerIdx >= 0) ? headerIdx + 1 : 0;
  const colMap = {};
  if (headerIdx >= 0) {
    const header = mat[headerIdx].map(c=>up(c));
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
    for (let c=0;c<7;c++) colMap[c] = c;
  }

  let weekDateCol = -1;
  if (headerIdx >= 0) {
    const headerRow = mat[headerIdx];
    for (let c=0;c<headerRow.length;c++){
      const cell = String(headerRow[c]||'').toLowerCase();
      if (cell.includes('week') || cell.includes('date')) { weekDateCol = c; break; }
    }
  }

  const rows = [];
  let pseudoCounter = 0;
  for (let r = startRow; r < mat.length; r++) {
    const row = mat[r];
    const weekAnchorRaw = (weekDateCol >= 0 && row[weekDateCol]) ? String(row[weekDateCol]) : null;
    let weekAnchor = null;
    if (weekAnchorRaw) {
      const d = dayjs(weekAnchorRaw);
      if (d.isValid()) weekAnchor = d;
    }
    for (const [colStr, wd] of Object.entries(colMap)) {
      const c = Number(colStr);
      const raw = (row[c] === undefined || row[c] === null) ? '' : row[c];
      if (isXXXorInvalid(raw)) continue;
      const num = extract3(raw);
      if (!num) continue;
      let date = null;
      if (weekAnchor) {
        date = weekAnchor.add(wd, 'day').startOf('day');
      } else {
        date = dayjs('1970-01-01').add(pseudoCounter, 'day');
      }
      rows.push({ date, num, weekday: wd, rowIndex: r, colIndex: c });
    }
    pseudoCounter += 7;
  }
  rows.sort((a,b) => a.date.valueOf() - b.date.valueOf());
  return rows;
}

// parse monthly grid: expects a DATE column then month columns (JAN, FEB...)
function parseMonthlyGrid(mat) {
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const headerIdx = mat.findIndex(row => row && row.some(c => up(c).includes('DATE')));
  const headerRow = mat[ headerIdx >= 0 ? headerIdx : 0 ].map(c => up(c));
  const monthMap = { 'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12 };
  const colToMonth = {};
  headerRow.forEach((cell, idx) => {
    if (!cell) return;
    for (const key of Object.keys(monthMap)) if (cell.includes(key)) { colToMonth[idx] = monthMap[key]; break; }
    const short = cell.slice(0,3);
    if (!colToMonth[idx] && monthMap[short]) colToMonth[idx] = monthMap[short];
  });
  let dateCol = headerRow.findIndex(c => up(c).includes('DATE'));
  if (dateCol < 0) dateCol = 0;
  const startRow = (headerIdx >= 0) ? headerIdx + 1 : 1;
  const rows = [];
  let pseudoCounter = 0;
  for (let r = startRow; r < mat.length; r++) {
    const row = mat[r];
    const dateRaw = row[dateCol];
    let dayNum = parseInt(String(dateRaw||'').trim(), 10);
    if (!dayNum || dayNum < 1 || dayNum > 31) { pseudoCounter++; continue; }
    for (const [cStr, mon] of Object.entries(colToMonth)) {
      const c = Number(cStr);
      const raw = row[c];
      if (isXXXorInvalid(raw)) continue;
      const num = extract3(raw);
      if (!num) continue;
      let date = null;
      const headerCell = headerRow[c] || '';
      const yMatch = headerCell.match(/20\d{2}/);
      if (yMatch) {
        date = dayjs(`${yMatch[0]}-${String(mon).padStart(2,'0')}-${String(dayNum).padStart(2,'0')}`);
      } else {
        for (let cc = 0; cc < Math.min(4, row.length); cc++) {
          const candidate = String(row[cc]||'');
          const y2 = candidate.match(/20\d{2}/);
          if (y2) { date = dayjs(`${y2[0]}-${String(mon).padStart(2,'0')}-${String(dayNum).padStart(2,'0')}`); break; }
        }
      }
      if (!date || !date.isValid()) date = dayjs('1970-01-01').add(pseudoCounter, 'day');
      rows.push({ date, num, day: dayNum, month: mon, r, c });
    }
    pseudoCounter++;
  }
  rows.sort((a,b) => a.date.valueOf() - b.date.valueOf());
  return rows;
}

// ---------------- build distributions ----------------
function buildDistributionFromCounter(counter, alphaTriplet = ALPHA_TRIPLET, alphaPos = ALPHA_POS, mix = MIX_TRIPLET) {
  const tripletDist = Object.create(null);
  const N = counter.total || 0;
  const denomTrip = (N + alphaTriplet * MAX_CLASSES);
  for (const t of TRIPLETS) {
    const cnt = counter.triplets[t] || 0;
    tripletDist[t] = (cnt + alphaTriplet) / denomTrip;
  }

  const denomPos = (N + alphaPos * 10);
  const p0 = {}, p1 = {}, p2 = {};
  for (let d=0; d<=9; d++) {
    const k = String(d);
    p0[k] = ((counter.pos0[k]||0) + alphaPos) / denomPos;
    p1[k] = ((counter.pos1[k]||0) + alphaPos) / denomPos;
    p2[k] = ((counter.pos2[k]||0) + alphaPos) / denomPos;
  }
  const posDist = Object.create(null);
  for (const a of Object.keys(p0)) for (const b of Object.keys(p1)) for (const c of Object.keys(p2)) {
    const t = `${a}${b}${c}`;
    posDist[t] = p0[a]*p1[b]*p2[c];
  }
  let sumPos = 0;
  for (const k of TRIPLETS) sumPos += (posDist[k]||0);
  if (sumPos <= 0) sumPos = 1;
  for (const k of TRIPLETS) posDist[k] = (posDist[k]||0) / sumPos;

  const final = Object.create(null);
  for (const k of TRIPLETS) final[k] = mix * tripletDist[k] + (1-mix) * posDist[k];

  let sumSq = 0;
  for (const k of TRIPLETS) { const p = final[k] || 0; sumSq += p*p; }
  return { dist: final, sumSquares: sumSq };
}

// ---------------- ranking helpers ----------------
function rankOfTrue(dist, trueTrip) {
  const arr = Object.entries(dist).sort((a,b) => b[1] - a[1]);
  for (let i=0;i<arr.length;i++) if (arr[i][0] === trueTrip) return i+1;
  return arr.length+1;
}
function isInTopK(dist, trueTrip, K) {
  const arr = Object.entries(dist).sort((a,b) => b[1] - a[1]).slice(0,K);
  return arr.some(x => x[0] === trueTrip);
}

// ---------------- expanding-window backtest ----------------
function runBacktest(records, mode='weekly') {
  records = records.slice().sort((a,b) => a.date.valueOf() - b.date.valueOf());
  const N = records.length;
  if (N <= MIN_TRAIN) {
    console.warn(`Not enough records for backtest in mode=${mode}. Records=${N}, MIN_TRAIN=${MIN_TRAIN}`);
    return null;
  }
  const groupCounters = {};
  const overallCounter = mkCounter();
  const results = [];

  for (let i=0;i<N;i++) {
    const rec = records[i];
    if (i < MIN_TRAIN) {
      addToCounter(overallCounter, rec.num, 1);
      let key = null;
      if (mode === 'weekly') key = String(rec.date.day());
      else if (mode === 'monthly') key = `${rec.date.month()+1}-${rec.date.date()}`;
      if (!groupCounters[key]) groupCounters[key] = mkCounter();
      addToCounter(groupCounters[key], rec.num, 1);
      continue;
    }

    let useCounter = null;
    if (mode === 'weekly') {
      const key = String(rec.date.day());
      useCounter = groupCounters[key] || mkCounter();
    } else {
      const mdKey = `${rec.date.month()+1}-${rec.date.date()}`;
      if (groupCounters[mdKey] && (groupCounters[mdKey].total || 0) > 0) useCounter = groupCounters[mdKey];
      else {
        const mKey = `M-${rec.date.month()+1}`;
        if (groupCounters[mKey] && (groupCounters[mKey].total || 0) > 0) useCounter = groupCounters[mKey];
        else useCounter = overallCounter;
      }
    }

    const { dist, sumSquares } = buildDistributionFromCounter(useCounter, ALPHA_TRIPLET, ALPHA_POS, MIX_TRIPLET);
    const p_true = dist[rec.num] || 0.0;
    const rank = rankOfTrue(dist, rec.num);
    const inTop = {};
    for (const K of TOP_KS) inTop[K] = isInTopK(dist, rec.num, K);
    const brier = sumSquares - 2*p_true + 1;

    results.push({
      date: rec.date.format('YYYY-MM-DD'),
      trueNum: rec.num,
      p_true,
      rank,
      brier,
      inTop,
      observationsUsed: useCounter.total || 0
    });

    addToCounter(overallCounter, rec.num, 1);
    if (mode === 'weekly') {
      const k = String(rec.date.day());
      if (!groupCounters[k]) groupCounters[k] = mkCounter();
      addToCounter(groupCounters[k], rec.num, 1);
    } else {
      const mdKey = `${rec.date.month()+1}-${rec.date.date()}`;
      if (!groupCounters[mdKey]) groupCounters[mdKey] = mkCounter();
      addToCounter(groupCounters[mdKey], rec.num, 1);
      const mKey = `M-${rec.date.month()+1}`;
      if (!groupCounters[mKey]) groupCounters[mKey] = mkCounter();
      addToCounter(groupCounters[mKey], rec.num, 1);
    }
  }

  const totalTests = results.length;
  const topKCounts = {}; TOP_KS.forEach(k => topKCounts[k] = 0);
  let mrrSum = 0; let brierSum = 0;
  for (const r of results) {
    for (const k of TOP_KS) if (r.inTop[k]) topKCounts[k] += 1;
    mrrSum += (1 / r.rank);
    brierSum += r.brier;
  }
  const topKRates = {}; TOP_KS.forEach(k => topKRates[k] = (topKCounts[k] / totalTests) || 0);
  const mrr = mrrSum / totalTests;
  const meanBrier = brierSum / totalTests;

  return { totalTests, topKRates, mrr, meanBrier, results };
}

// ---------------- high-level parse logic ----------------
async function parseWeeklyFile(path) {
  const mat = await sheetToMatrix(path, ['weekly','Week','Sheet1','Sheet']);
  const flat = tryParseFlattened(mat);
  if (flat && flat.length > 0) {
    let pseudo = 0;
    for (const r of flat) if (!r.date) { r.date = dayjs('1970-01-01').add(pseudo, 'day'); pseudo++; }
    return flat.map(r => ({ date: r.date, num: r.num }));
  }
  const grid = parseWeeklyGrid(mat);
  return grid.map(g => ({ date: g.date, num: g.num }));
}

async function parseMonthlyFile(path) {
  const mat = await sheetToMatrix(path, ['monthly','Month','Sheet1','Sheet']);
  const flat = tryParseFlattened(mat);
  if (flat && flat.length > 0) {
    let pseudo = 0;
    for (const r of flat) if (!r.date) { r.date = dayjs('1970-01-01').add(pseudo,'day'); pseudo++; }
    return flat.map(r => ({ date: r.date, num: r.num }));
  }
  const grid = parseMonthlyGrid(mat);
  return grid.map(g => ({ date: g.date, num: g.num }));
}

// ---------------- write CSV results helper (writes into outputFolder) ----------------
function writeCsv(rows, outPath) {
  if (!rows || rows.length === 0) return;
  const header = Object.keys(rows[0]);
  const lines = [header.join(',')];
  for (const r of rows) {
    const vals = header.map(h => {
      const v = r[h];
      if (v === null || v === undefined) return '';
      if (typeof v === 'object') return `"${JSON.stringify(v).replace(/"/g,'""')}"`;
      return String(v).replace(/"/g,'""');
    });
    lines.push(vals.join(','));
  }
  fs.writeFileSync(outPath, lines.join('\n'), 'utf8');
}

// ---------------- main ----------------
(async () => {
  const args = process.argv.slice(2);
  if (args.length < 2) {
    console.error('Usage: node backtest_lottery.js weekly.xlsx monthly.xlsx');
    process.exit(1);
  }
  const weeklyPath = args[0];
  const monthlyPath = args[1];

  console.log('Parsing weekly file:', weeklyPath);
  const weeklyRecords = await parseWeeklyFile(weeklyPath);
  console.log('Parsed weekly records:', weeklyRecords.length);

  console.log('Parsing monthly file:', monthlyPath);
  const monthlyRecords = await parseMonthlyFile(monthlyPath);
  console.log('Parsed monthly records:', monthlyRecords.length);

  console.log('\nRunning weekly backtest (separate)...');
  const wres = runBacktest(weeklyRecords, 'weekly');
  if (wres) {
    console.log('Weekly tests:', wres.totalTests);
    console.log('Weekly Top-K rates:', wres.topKRates);
    console.log('Weekly MRR:', wres.mrr.toFixed(6));
    console.log('Weekly Mean Brier:', wres.meanBrier.toFixed(6));
    const wrows = wres.results.map(r => ({
      date: r.date, trueNum: r.trueNum, p_true: r.p_true, rank: r.rank, brier: r.brier, observationsUsed: r.observationsUsed
    }));
    const wOut = path.join(outputFolder, 'weekly_backtest_results.csv');
    writeCsv(wrows, wOut);
    console.log('Wrote', wOut);
  }

  console.log('\nRunning monthly backtest (separate)...');
  const mres = runBacktest(monthlyRecords, 'monthly');
  if (mres) {
    console.log('Monthly tests:', mres.totalTests);
    console.log('Monthly Top-K rates:', mres.topKRates);
    console.log('Monthly MRR:', mres.mrr.toFixed(6));
    console.log('Monthly Mean Brier:', mres.meanBrier.toFixed(6));
    const mrows = mres.results.map(r => ({
      date: r.date, trueNum: r.trueNum, p_true: r.p_true, rank: r.rank, brier: r.brier, observationsUsed: r.observationsUsed
    }));
    const mOut = path.join(outputFolder, 'monthly_backtest_results.csv');
    writeCsv(mrows, mOut);
    console.log('Wrote', mOut);
  }

})();
