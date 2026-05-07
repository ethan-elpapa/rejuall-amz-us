// Vercel Serverless Function — Google Sheets 데이터를 안전하게 프록시
// 인증: 서비스 계정 (환경변수 GOOGLE_SERVICE_ACCOUNT_JSON)
// 캐시: Vercel Edge 60초 (s-maxage)

const { google } = require('googleapis');

const SPREADSHEET_ID = '10d21g2iUkqb2uRVEw9ZCSQgmOAqKVyuIiOSNBozl1ok';

const PRODUCT_NAMES = {
  'B0FGDR67R5': 'PDRN Cream (구)',
  'B0FN7LDTB7': 'PDRN Cream Max',
  'B0FPDWZ5X2': 'Ceramide Cream',
  'B0FP4WJXD5': 'Retino-Mela Serum',
  'B0FN7L65C1': 'PDRN Cream',
  'B0GL22N4GY': 'PDRN Lip Serum',
  'B0GL26G5G2': 'PDRN Mask',
  'B0GMWJBZFF': 'Skincare Program',
  'B0FPXC3P3C': 'PDRN Parent'
};

const PRODUCT_COLORS = {
  'B0FGDR67R5': '#6c8ef7',
  'B0FN7LDTB7': '#10b981',
  'B0FPDWZ5X2': '#f59e0b',
  'B0FP4WJXD5': '#ec4899',
  'B0FN7L65C1': '#8b5cf6',
  'B0GL22N4GY': '#14b8a6',
  'B0GL26G5G2': '#f97316',
  'B0GMWJBZFF': '#06b6d4',
  'B0FPXC3P3C': '#a3e635'
};

// ── 헬퍼 ──────────────────────────────────────────────
function normalizeHeader(s) {
  return String(s || '').toLowerCase()
    .replace(/[()[\]{}]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function findColIdx(headers, keyword) {
  const kw = normalizeHeader(keyword);
  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === kw) return i;
  }
  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]).indexOf(kw) > -1) return i;
  }
  return -1;
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  return parseFloat(String(v).replace(/[$,%\s₩]/g, '')) || 0;
}

function normalizeDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return v.toISOString().substring(0, 10);
  }
  if (typeof v === 'number') {
    const d = new Date((v - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d.toISOString().substring(0, 10);
  }
  const s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  const p = new Date(s);
  if (!isNaN(p.getTime())) return p.toISOString().substring(0, 10);
  return '';
}

function parseMonthKey(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return v.toISOString().substring(0, 7);
  if (typeof v === 'number') {
    if (v >= 1 && v <= 12 && v === Math.floor(v)) return String(v);
    const d = new Date((v - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d.toISOString().substring(0, 7);
  }
  const s = String(v).trim();
  if (!s) return '';
  let m = s.match(/(\d{4})[년\-/.\s]+(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/^(\d{1,2})\s*월?$/);
  if (m) {
    const n = parseInt(m[1], 10);
    if (n >= 1 && n <= 12) return String(n);
  }
  return '';
}

// ── 시트 선택 로직 ─────────────────────────────────────
async function getSheetMeta(sheets) {
  const meta = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'sheets(properties(title,gridProperties(rowCount,columnCount)))'
  });
  return meta.data.sheets.map(s => s.properties);
}

function pickRawSheet(allSheets, allValues) {
  // Date + Sessions - Total + Child ASIN 헤더가 모두 있는 시트 중 행이 가장 많은 것
  let best = null;
  for (const sh of allSheets) {
    const vals = allValues[sh.title];
    if (!vals || vals.length < 2) continue;
    const header = vals[0].map(h => String(h).toLowerCase());
    const hasDate = header.some(h => h === 'date');
    const hasSessions = header.some(h => h.indexOf('sessions - total') > -1);
    const hasChild = header.some(h => h.indexOf('child asin') > -1);
    if (hasDate && hasSessions && hasChild) {
      if (!best || vals.length > allValues[best.title].length) best = sh;
    }
  }
  return best;
}

function pickGoalSheet(allSheets) {
  for (const sh of allSheets) {
    const name = sh.title.replace(/\s/g, '');
    if (name.indexOf('목표') > -1 || name.toLowerCase().indexOf('goal') > -1 || name.toLowerCase().indexOf('target') > -1) {
      return sh;
    }
  }
  return null;
}

function pickInfoSheet(allSheets) {
  for (const sh of allSheets) {
    const upper = sh.title.replace(/\s/g, '').toUpperCase();
    if (upper === 'ASIN-INFO' || upper === 'ASININFO' || upper === 'PRODUCT' || upper === '제품') return sh;
  }
  return null;
}

// ── 파서 ───────────────────────────────────────────────
function parseRawRows(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const idx = {
    month:        findColIdx(headers, 'month'),
    week:         findColIdx(headers, 'week'),
    weekCode:     findColIdx(headers, '월-주차'),
    date:         findColIdx(headers, 'date'),
    parent:       findColIdx(headers, 'parent asin'),
    child:        findColIdx(headers, 'child asin'),
    sessions:     findColIdx(headers, 'sessions - total'),
    sessionsB2B:  findColIdx(headers, 'sessions - total - b2b'),
    pageViews:    findColIdx(headers, 'page views - total'),
    pageViewsB2B: findColIdx(headers, 'page views - total - b2b'),
    buyBox:       findColIdx(headers, 'featured offer (buy box) percentage'),
    units:        findColIdx(headers, 'units ordered'),
    unitsB2B:     findColIdx(headers, 'units ordered - b2b'),
    cvr:          findColIdx(headers, 'unit session percentage'),
    sales:        findColIdx(headers, 'ordered product sales'),
    salesB2B:     findColIdx(headers, 'ordered product sales - b2b'),
    orderItems:   findColIdx(headers, 'total order items')
  };

  const cellStr = (r, i) => (i < 0 || i >= r.length) ? '' : String(r[i] ?? '').trim();
  const cellNum = (r, i) => (i < 0 || i >= r.length) ? 0  : toNum(r[i]);

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const dateStr = normalizeDate(idx.date >= 0 ? r[idx.date] : null);
    if (!dateStr) continue;
    const childAsin = cellStr(r, idx.child);
    if (!childAsin) continue;

    out.push({
      month: cellStr(r, idx.month),
      week: cellStr(r, idx.week),
      weekCode: cellStr(r, idx.weekCode),
      date: dateStr,
      parentAsin: cellStr(r, idx.parent),
      childAsin: childAsin,
      sessions: cellNum(r, idx.sessions),
      sessionsB2B: cellNum(r, idx.sessionsB2B),
      pageViews: cellNum(r, idx.pageViews),
      pageViewsB2B: cellNum(r, idx.pageViewsB2B),
      buyBox: cellNum(r, idx.buyBox),
      units: cellNum(r, idx.units),
      unitsB2B: cellNum(r, idx.unitsB2B),
      cvr: cellNum(r, idx.cvr) * 100,
      sales: cellNum(r, idx.sales),
      salesB2B: cellNum(r, idx.salesB2B),
      orderItems: cellNum(r, idx.orderItems)
    });
  }
  return out;
}

function parseInfoSheet(values) {
  const map = {};
  const knownAsins = [];
  if (!values || values.length < 2) return { map, knownAsins };
  for (let i = 1; i < values.length; i++) {
    const row = values[i] || [];
    const asin = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (asin.indexOf('B0') === 0 && name) {
      map[asin] = name;
      knownAsins.push(asin);
    }
  }
  return { map, knownAsins };
}

function parseGoalSheet(values) {
  const goals = {};
  if (!values || values.length < 1) return goals;

  // 케이스 1: 가로형 — 어느 행이 헤더(월키 3개+)인지 찾고, 다음 행을 값으로
  for (let rowH = 0; rowH < Math.min(3, values.length); rowH++) {
    const monthCols = [];
    const row = values[rowH] || [];
    for (let c = 0; c < row.length; c++) {
      const key = parseMonthKey(row[c]);
      if (key) monthCols.push({ col: c, key });
    }
    if (monthCols.length >= 3 && rowH + 1 < values.length) {
      let found = 0;
      for (const mc of monthCols) {
        const v = toNum((values[rowH + 1] || [])[mc.col]);
        if (v) { goals[mc.key] = v; found++; }
      }
      if (found >= 3) return goals;
    }
  }

  // 케이스 2: 세로형 — 가장 많이 매칭되는 (월컬럼, 값컬럼) 페어
  let bestPair = null, bestCount = 0;
  const maxCol = Math.max(...values.map(r => r.length));
  for (let mc = 0; mc < maxCol; mc++) {
    for (let vc = 0; vc < maxCol; vc++) {
      if (mc === vc) continue;
      const tmp = {};
      let cnt = 0;
      for (let r = 0; r < values.length; r++) {
        const row = values[r] || [];
        const key = parseMonthKey(row[mc]);
        const val = toNum(row[vc]);
        if (key && val) { tmp[key] = val; cnt++; }
      }
      if (cnt > bestCount) { bestCount = cnt; bestPair = tmp; }
    }
  }
  return bestPair || {};
}

// ── 메인 핸들러 ────────────────────────────────────────
let cached = null; // 메모리 캐시 (인스턴스 살아있는 동안)

module.exports = async (req, res) => {
  try {
    const fresh = req.query && req.query.fresh === '1';
    const now = Date.now();
    if (!fresh && cached && (now - cached.at) < 60_000) {
      res.setHeader('Cache-Control', 's-maxage=60, stale-while-revalidate=300');
      res.setHeader('X-Cache', 'HIT');
      return res.status(200).json(cached.data);
    }

    const credsRaw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
    if (!credsRaw) {
      return res.status(500).json({ ok: false, error: 'GOOGLE_SERVICE_ACCOUNT_JSON 환경변수 미설정' });
    }
    const credentials = JSON.parse(credsRaw);

    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly']
    });
    const sheets = google.sheets({ version: 'v4', auth });

    const sheetProps = await getSheetMeta(sheets);

    // 모든 시트 데이터 한 번에 불러오기 (batchGet)
    const ranges = sheetProps.map(s => `'${s.title.replace(/'/g, "''")}'`);
    const batch = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges,
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });

    const allValues = {};
    batch.data.valueRanges.forEach((vr, i) => {
      allValues[sheetProps[i].title] = vr.values || [];
    });

    const rawSheet  = pickRawSheet(sheetProps, allValues);
    const infoSheet = pickInfoSheet(sheetProps);
    const goalSheet = pickGoalSheet(sheetProps);

    const rows         = rawSheet  ? parseRawRows(allValues[rawSheet.title]) : [];
    const productInfo  = infoSheet ? parseInfoSheet(allValues[infoSheet.title]) : { map: {}, knownAsins: [] };
    const monthlyGoals = goalSheet ? parseGoalSheet(allValues[goalSheet.title]) : {};

    // 폴백: ASIN-INFO 비어있으면 PRODUCT_NAMES 사용
    if (productInfo.knownAsins.length === 0) {
      for (const asin of Object.keys(PRODUCT_NAMES)) {
        productInfo.map[asin] = PRODUCT_NAMES[asin];
        productInfo.knownAsins.push(asin);
      }
    }

    const data = {
      ok: true,
      rows,
      productMap: productInfo.map,
      knownAsins: productInfo.knownAsins,
      productColors: PRODUCT_COLORS,
      monthlyGoals,
      updatedAt: new Date().toISOString(),
      diag: {
        rawSheet: rawSheet ? rawSheet.title : null,
        rowCount: rows.length,
        goalSheet: goalSheet ? goalSheet.title : null,
        goalKeys: Object.keys(monthlyGoals).length
      }
    };

    cached = { data, at: now };

    res.setHeader('Cache-Control', 's-maxage=60, stale-while-revalidate=300');
    res.setHeader('X-Cache', 'MISS');
    res.status(200).json(data);
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message || String(e), stack: e.stack });
  }
};
