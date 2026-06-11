// Vercel Serverless Function — Google Sheets 데이터를 안전하게 프록시
// 인증: 서비스 계정 (환경변수 GOOGLE_SERVICE_ACCOUNT_JSON)
// 캐시: Vercel Edge 60초 (s-maxage)

const { google } = require('googleapis');

const SPREADSHEET_ID = '10d21g2iUkqb2uRVEw9ZCSQgmOAqKVyuIiOSNBozl1ok';

const PRODUCT_NAMES = {
  'B0FGDR67R5': 'PDRN 20ml (구)',
  'B0FN7LDTB7': 'PDRN MAX',
  'B0FPDWZ5X2': '세라마이드',
  'B0FP4WJXD5': '레티노멜라',
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
    .replace(/[\u2010-\u2015\u2212]/g, '-')   // em/en/figure dashes → hyphen
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
  // Date + Sessions + Child ASIN 헤더가 모두 있는 시트 중 행이 가장 많은 것
  let best = null;
  for (const sh of allSheets) {
    const vals = allValues[sh.title];
    if (!vals || vals.length < 2) continue;
    const header = (vals[0] || []).map(h => normalizeHeader(h));
    const hasDate     = header.some(h => h === 'date' || h === '날짜');
    const hasSessions = header.some(h => h.indexOf('session') > -1);
    const hasChild    = header.some(h => h.indexOf('child asin') > -1 || h.indexOf('child') > -1 && h.indexOf('asin') > -1);
    if (hasDate && hasSessions && hasChild) {
      if (!best || vals.length > allValues[best.title].length) best = sh;
    }
  }
  // 폴백: 가장 행 많은 시트 중 child asin이 있는 것
  if (!best) {
    for (const sh of allSheets) {
      const vals = allValues[sh.title];
      if (!vals || vals.length < 2) continue;
      const header = (vals[0] || []).map(h => normalizeHeader(h));
      if (header.some(h => h.indexOf('child') > -1 && h.indexOf('asin') > -1)) {
        if (!best || vals.length > allValues[best.title].length) best = sh;
      }
    }
  }
  return best;
}

function pickGoalSheet(allSheets) {
  // 1순위: 정확히 '목표 매출' 또는 '목표' 포함 (target/goal는 너무 광범위해서 후순위)
  for (const sh of allSheets) {
    const name = sh.title.replace(/\s/g, '');
    if (name.indexOf('목표') > -1) return sh;
  }
  // 2순위: 영문 키워드 — 단, 'targeting' 같이 다른 의미는 제외
  for (const sh of allSheets) {
    const lower = sh.title.toLowerCase();
    if (lower.indexOf('targeting') > -1) continue; // 광고 타겟팅 리포트 제외
    if (/(^|[\s_-])goal($|[\s_-])/.test(lower) || /(^|[\s_-])target($|[\s_-])/.test(lower) || lower.indexOf('monthly goal') > -1) {
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

function pickAdRawSheet(allSheets) {
  for (const sh of allSheets) {
    const name = sh.title.replace(/\s/g, '');
    if (name.indexOf('내부광고RAW') > -1 || name.indexOf('광고RAW') > -1 || /ad.*raw/i.test(sh.title)) return sh;
  }
  return null;
}

function pickSearchTermSheet(allSheets) {
  for (const sh of allSheets) {
    const lower = sh.title.toLowerCase().replace(/\s/g, '');
    if (lower.indexOf('searchterm') > -1 || sh.title.indexOf('서치텀') > -1) return sh;
  }
  return null;
}

function pickTargetingSheet(allSheets) {
  for (const sh of allSheets) {
    if (sh.title.toLowerCase().indexOf('targeting') > -1) return sh;
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

// 내부 광고 RAW 파서
function parseAdRaw(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const idx = {
    week:        findColIdx(headers, '주차'),
    date:        findColIdx(headers, 'date'),
    asin:        findColIdx(headers, 'asin'),
    portfolio:   findColIdx(headers, 'portfolio'),
    impressions: findColIdx(headers, 'impressions'),
    clicks:      findColIdx(headers, 'clicks'),
    ctr:         findColIdx(headers, 'ctr'),
    spend:       findColIdx(headers, 'spend'),
    cpc:         findColIdx(headers, 'cpc'),
    orders:      findColIdx(headers, 'orders'),
    sales:       findColIdx(headers, 'sales'),
    acos:        findColIdx(headers, 'acos'),
    roas:        findColIdx(headers, 'roas'),
    ntbOrders:   findColIdx(headers, 'ntb orders')
  };
  const cellStr = (r, i) => (i < 0 || i >= r.length) ? '' : String(r[i] ?? '').trim();
  const cellNum = (r, i) => (i < 0 || i >= r.length) ? 0  : toNum(r[i]);

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const dateStr = normalizeDate(idx.date >= 0 ? r[idx.date] : null);
    if (!dateStr) continue;
    out.push({
      week:        cellStr(r, idx.week),
      date:        dateStr,
      asin:        cellStr(r, idx.asin),
      portfolio:   cellStr(r, idx.portfolio),
      impressions: cellNum(r, idx.impressions),
      clicks:      cellNum(r, idx.clicks),
      ctr:         cellNum(r, idx.ctr),
      spend:       cellNum(r, idx.spend),
      cpc:         cellNum(r, idx.cpc),
      orders:      cellNum(r, idx.orders),
      sales:       cellNum(r, idx.sales),
      acos:        cellNum(r, idx.acos),
      roas:        cellNum(r, idx.roas),
      ntbOrders:   cellNum(r, idx.ntbOrders)
    });
  }
  return out;
}

// Search Term Report 파서 (term + portfolio + matchType 단위 집계 — 응답 크기 축소)
function parseSearchTerms(values) {
  if (!values || values.length < 2) return { agg: [], byWeek: [] };
  const headers = values[0];
  const idx = {
    week:       findColIdx(headers, '주차'),
    portfolio:  findColIdx(headers, 'portfolio name'),
    campaign:   findColIdx(headers, 'campaign name'),
    matchType:  findColIdx(headers, 'match type'),
    term:       findColIdx(headers, 'customer search term'),
    impressions:findColIdx(headers, 'impressions'),
    clicks:     findColIdx(headers, 'clicks'),
    spend:      findColIdx(headers, 'spend'),
    sales:      findColIdx(headers, '7 day total sales'),
    orders:     findColIdx(headers, 'orders')
  };
  const cellStr = (r, i) => (i < 0 || i >= r.length) ? '' : String(r[i] ?? '').trim();
  const cellNum = (r, i) => (i < 0 || i >= r.length) ? 0  : toNum(r[i]);

  // term × portfolio × matchType × week 원장
  const keyWeekMap = {};
  const weekMap = {};
  const allWeeks = new Set();
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const term = cellStr(r, idx.term);
    if (!term) continue;
    const portfolio = cellStr(r, idx.portfolio);
    const matchType = cellStr(r, idx.matchType);
    const campaign = cellStr(r, idx.campaign);
    const week = cellStr(r, idx.week);
    if (week) allWeeks.add(week);
    const impressions = cellNum(r, idx.impressions);
    const clicks = cellNum(r, idx.clicks);
    const spend = cellNum(r, idx.spend);
    const sales = cellNum(r, idx.sales);
    const orders = cellNum(r, idx.orders);

    const key = term + '|' + portfolio + '|' + matchType;
    if (!keyWeekMap[key]) {
      keyWeekMap[key] = { meta: { term, portfolio, matchType, campaign }, weeks: {} };
    }
    const km = keyWeekMap[key];
    if (!km.weeks[week]) km.weeks[week] = { i:0, c:0, s:0, sa:0, o:0 };
    const ww = km.weeks[week];
    ww.i += impressions; ww.c += clicks; ww.s += spend; ww.sa += sales; ww.o += orders;

    // week × portfolio 집계 (트렌드용)
    const wkey = week + '|' + portfolio;
    if (!weekMap[wkey]) {
      weekMap[wkey] = { week, portfolio, impressions: 0, clicks: 0, spend: 0, sales: 0, orders: 0, termCount: 0 };
    }
    const wagg = weekMap[wkey];
    wagg.impressions += impressions;
    wagg.clicks      += clicks;
    wagg.spend       += spend;
    wagg.sales       += sales;
    wagg.orders      += orders;
    wagg.termCount   += 1;
  }

  // 주차 라벨 정렬 (예: W12-1 < W1-1 < W2-1 ... W4-3)
  // W12를 0, W1=1, ..., W11=11로 매핑 (현재 데이터 스냅샷 기준)
  function weekRank(w) {
    const m = String(w).match(/W(\d+)-(\d+)/);
    if (!m) return [99, 99];
    const mo = parseInt(m[1], 10);
    const sub = parseInt(m[2], 10);
    return [mo === 12 ? 0 : mo, sub];
  }
  const weeksSorted = Array.from(allWeeks).sort((a, b) => {
    const ra = weekRank(a), rb = weekRank(b);
    return ra[0] - rb[0] || ra[1] - rb[1];
  });

  function buildAgg(weekSet /* Set | null = all */) {
    const out = [];
    for (const key in keyWeekMap) {
      const km = keyWeekMap[key];
      let imp=0, clk=0, sp=0, sa=0, od=0, wc=0;
      for (const w in km.weeks) {
        if (weekSet && !weekSet.has(w)) continue;
        const ww = km.weeks[w];
        imp += ww.i; clk += ww.c; sp += ww.s; sa += ww.sa; od += ww.o;
        wc++;
      }
      if (imp + clk + sp + sa + od === 0) continue;
      out.push({
        term: km.meta.term,
        portfolio: km.meta.portfolio,
        matchType: km.meta.matchType,
        campaign: km.meta.campaign,
        impressions: imp,
        clicks: clk,
        spend: +sp.toFixed(2),
        sales: +sa.toFixed(2),
        orders: od,
        ctr:  imp > 0 ? clk / imp : 0,
        cpc:  clk > 0 ? sp / clk : 0,
        cvr:  clk > 0 ? od / clk : 0,
        acos: sa > 0 ? sp / sa : 0,
        roas: sp > 0 ? sa / sp : 0,
        weekCount: wc
      });
    }
    return out;
  }

  const lastN = n => new Set(weeksSorted.slice(-n));
  const agg     = buildAgg(null);
  const last1w  = buildAgg(lastN(1));
  const last4w  = buildAgg(lastN(4));
  const last12w = buildAgg(lastN(12));

  const byWeek = Object.values(weekMap).map(w => ({
    week: w.week,
    portfolio: w.portfolio,
    impressions: w.impressions,
    clicks: w.clicks,
    spend: +w.spend.toFixed(2),
    sales: +w.sales.toFixed(2),
    orders: w.orders,
    termCount: w.termCount,
    acos: w.sales > 0 ? w.spend / w.sales : 0,
    roas: w.spend > 0 ? w.sales / w.spend : 0
  }));

  // term × portfolio × matchType × week 플랫 배열 (키워드 검색용)
  const byTermWeek = [];
  for (const key in keyWeekMap) {
    const km = keyWeekMap[key];
    for (const w in km.weeks) {
      const ww = km.weeks[w];
      // 무의미한 0행 제외
      if (!(ww.i || ww.c || ww.s || ww.sa || ww.o)) continue;
      byTermWeek.push({
        t:  km.meta.term,
        p:  km.meta.portfolio,
        mt: km.meta.matchType,
        w:  w,
        i:  ww.i,
        c:  ww.c,
        s:  +ww.s.toFixed(2),
        sa: +ww.sa.toFixed(2),
        o:  ww.o
      });
    }
  }

  return { agg, byWeek, periods: { all: agg, last1w, last4w, last12w }, weeksSorted, byTermWeek };
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

    const rawSheet     = pickRawSheet(sheetProps, allValues);
    const infoSheet    = pickInfoSheet(sheetProps);
    const goalSheet    = pickGoalSheet(sheetProps);
    const adRawSheet   = pickAdRawSheet(sheetProps);
    const stSheet      = pickSearchTermSheet(sheetProps);

    const rows         = rawSheet    ? parseRawRows(allValues[rawSheet.title]) : [];
    const productInfo  = infoSheet   ? parseInfoSheet(allValues[infoSheet.title]) : { map: {}, knownAsins: [] };
    const monthlyGoals = goalSheet   ? parseGoalSheet(allValues[goalSheet.title]) : {};
    // 2025-10-18 이전 데이터는 포트폴리오 미설정 + Sales 컬럼 손상으로 제외
    const AD_DATA_START = '2025-10-18';
    const adRowsAll    = adRawSheet  ? parseAdRaw(allValues[adRawSheet.title]) : [];
    const adRows       = adRowsAll.filter(r => r.date && r.date >= AD_DATA_START && r.portfolio);
    const stParsed     = stSheet     ? parseSearchTerms(allValues[stSheet.title]) : { agg: [], byWeek: [] };

    // PRODUCT_NAMES 우선 적용 (시트 값보다 우선)
    for (const asin of Object.keys(PRODUCT_NAMES)) {
      productInfo.map[asin] = PRODUCT_NAMES[asin];
      if (productInfo.knownAsins.indexOf(asin) === -1) {
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
      adRows,
      searchTerms: stParsed.agg,
      searchTermsWeekly: stParsed.byWeek,
      searchTermsPeriods: stParsed.periods,
      searchTermsWeeks: stParsed.weeksSorted,
      searchTermsByTermWeek: stParsed.byTermWeek,
      updatedAt: new Date().toISOString(),
      diag: {
        rawSheet: rawSheet ? rawSheet.title : null,
        rowCount: rows.length,
        goalSheet: goalSheet ? goalSheet.title : null,
        goalKeys: Object.keys(monthlyGoals).length,
        adSheet: adRawSheet ? adRawSheet.title : null,
        adRowCount: adRows.length,
        stSheet: stSheet ? stSheet.title : null,
        stAggCount: stParsed.agg.length,
        stWeekCount: stParsed.byWeek.length
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
