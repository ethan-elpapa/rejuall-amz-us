// Vercel Serverless Function — Search Query Performance 데이터 프록시
// Source: Google Sheets `US_Search Query Performance_RAW`
// 인증: 서비스 계정 (환경변수 GOOGLE_SERVICE_ACCOUNT_JSON)
// 캐시: Vercel Edge 5분 (s-maxage=300)

const { google } = require('googleapis');

const SQP_SPREADSHEET_ID = '1msAFnZjSfNXnxD_wh49S2i2Ukf3zO_KcGqHeN0jmNLU';

function normalizeHeader(s) {
  return String(s || '').toLowerCase()
    .replace(/[()[\]{}]/g, '')
    .replace(/[\u2010-\u2015\u2212]/g, '-')
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
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).replace(/[$,%\s₩]/g, '');
  if (s === '') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseSqpRows(values) {
  if (!values || values.length < 2) return [];
  const headers = values[0];
  const idx = {
    weekCode:  findColIdx(headers, '주차'),
    weekNo:    findColIdx(headers, 'week no'),
    query:     findColIdx(headers, 'search query'),
    score:     findColIdx(headers, 'search query score'),
    volume:    findColIdx(headers, 'search query volume'),
    delta:     findColIdx(headers, '+%')
  };
  // delta header may render as '+%', '％', or similar — fallback scan
  if (idx.delta < 0) {
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (h === '+%' || h === '%' || h.indexOf('%') > -1) { idx.delta = i; break; }
    }
  }

  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i] || [];
    const weekCode = String(r[idx.weekCode] ?? '').trim();
    const query    = String(r[idx.query] ?? '').trim();
    if (!weekCode || !query) continue;

    out.push({
      weekCode: weekCode,
      weekNo:   String(r[idx.weekNo] ?? '').trim(),
      query:    query,
      score:    toNum(r[idx.score]),
      volume:   toNum(r[idx.volume]),
      delta:    toNum(r[idx.delta])
    });
  }
  return out;
}

let cached = null;

module.exports = async (req, res) => {
  try {
    const fresh = req.query && req.query.fresh === '1';
    const now = Date.now();
    if (!fresh && cached && (now - cached.at) < 300_000) {
      res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=600');
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

    const meta = await sheets.spreadsheets.get({
      spreadsheetId: SQP_SPREADSHEET_ID,
      fields: 'sheets(properties(title,sheetId))'
    });
    const sheetProps = meta.data.sheets.map(s => s.properties);

    const ranges = sheetProps.map(s => `'${s.title.replace(/'/g, "''")}'`);
    const batch = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SQP_SPREADSHEET_ID,
      ranges,
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING'
    });

    // 모든 시트의 데이터 병합 — 헤더에 'search query' 포함된 시트만 채택
    const allRows = [];
    const sheetDiag = [];
    batch.data.valueRanges.forEach((vr, i) => {
      const title = sheetProps[i].title;
      const vals = vr.values || [];
      const header = (vals[0] || []).map(h => normalizeHeader(h));
      const looksLikeData = header.some(h => h.indexOf('search query') > -1)
                         && header.some(h => h.indexOf('주차') > -1 || h.indexOf('week') > -1);
      sheetDiag.push({ title, rows: vals.length, looksLikeData });
      if (looksLikeData) {
        const parsed = parseSqpRows(vals);
        for (const r of parsed) allRows.push(Object.assign({ sheet: title }, r));
      }
    });

    // 주차 목록 (정렬): Week 번호 기준 오름차순
    const weekMap = {};
    for (const r of allRows) {
      const key = r.weekCode;
      if (!key) continue;
      if (!weekMap[key]) {
        const num = (r.weekNo || '').match(/(\d+)/);
        weekMap[key] = { code: key, label: r.weekNo || key, num: num ? parseInt(num[1], 10) : 999 };
      }
    }
    const weeks = Object.values(weekMap).sort((a, b) => a.num - b.num);

    const data = {
      ok: true,
      rows: allRows,
      weeks,
      updatedAt: new Date().toISOString(),
      diag: {
        sheets: sheetDiag,
        rowCount: allRows.length
      }
    };

    cached = { data, at: now };

    res.setHeader('Cache-Control', 's-maxage=300, stale-while-revalidate=600');
    res.setHeader('X-Cache', 'MISS');
    res.status(200).json(data);
  } catch (e) {
    res.status(500).json({ ok: false, error: e.message || String(e), stack: e.stack });
  }
};
