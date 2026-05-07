// ============================================================
// Dr. Reju-All Amazon Dashboard — Google Apps Script Backend
// 헤더 이름 기반 자동 매핑 (열 추가/삭제에 안전)
// ============================================================

var SPREADSHEET_ID = '10d21g2iUkqb2uRVEw9ZCSQgmOAqKVyuIiOSNBozl1ok';

var PRODUCT_NAMES = {
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

var PRODUCT_COLORS = {
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

function doGet(e) {
  // ?api=1 또는 ?format=json 으로 호출하면 JSON 반환 (Vercel 등 외부 호스팅용)
  if (e && e.parameter && (e.parameter.api === '1' || e.parameter.format === 'json')) {
    var json = getCachedJson_(e.parameter.fresh === '1');
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
  // 기본: HTML (Apps Script 내장 호스팅용)
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dr. Reju-All Amazon Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 5분 캐시 — 시트 읽기 비용 회피 (?fresh=1 로 강제 재계산 가능)
function getCachedJson_(forceFresh) {
  var cache = CacheService.getScriptCache();
  var KEY = 'dashboard_v2';
  if (!forceFresh) {
    var hit = readChunked_(cache, KEY);
    if (hit) return hit;
  }
  var data = getDashboardData();
  data._cachedAt = new Date().toISOString();
  var json = JSON.stringify(data);
  writeChunked_(cache, KEY, json, 300); // 5분
  return json;
}

function writeChunked_(cache, key, str, ttl) {
  var size = 90000; // 셀당 100KB 제한
  var chunks = Math.ceil(str.length / size);
  var pairs = {};
  pairs[key + '_meta'] = String(chunks);
  for (var i = 0; i < chunks; i++) {
    pairs[key + '_' + i] = str.substring(i * size, (i + 1) * size);
  }
  try { cache.putAll(pairs, ttl); } catch(e){}
}

function readChunked_(cache, key) {
  var meta = cache.get(key + '_meta');
  if (!meta) return null;
  var chunks = parseInt(meta, 10);
  if (!chunks) return null;
  var keys = [];
  for (var i = 0; i < chunks; i++) keys.push(key + '_' + i);
  var got = cache.getAll(keys);
  var out = '';
  for (var j = 0; j < chunks; j++) {
    var part = got[key + '_' + j];
    if (!part) return null; // 일부 만료시 무효
    out += part;
  }
  return out;
}

function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = findRawDataSheet(ss);
    var rows = readRawData(sheet);
    var productInfo = buildProductMap(ss);

    var goals = readMonthlyGoals(ss);

    return {
      ok: true,
      rows: rows,
      productMap:    productInfo.map,
      knownAsins:    productInfo.knownAsins,
      productColors: PRODUCT_COLORS,
      monthlyGoals:  goals,
      updatedAt:     new Date().toISOString(),
      diag: {
        sheetName: sheet ? sheet.getName() : null,
        rowCount:  rows.length
      }
    };
  } catch (e) {
    return { ok: false, error: e.toString() + '\n' + (e.stack || '') };
  }
}

// ── 헤더 인덱스 찾기 (괄호/특수문자 무시) ──────────────────
function normalizeHeader_(s) {
  return String(s).toLowerCase()
    .replace(/[()[\]{}]/g, '')   // 괄호 제거: "(Parent) ASIN" → "parent asin"
    .replace(/\s+/g, ' ')        // 다중 공백 정규화
    .trim();
}

function findColIdx(headers, keyword) {
  var kw = normalizeHeader_(keyword);
  // 1순위: 정확 일치
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader_(headers[i]) === kw) return i;
  }
  // 2순위: 부분 일치
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader_(headers[i]).indexOf(kw) > -1) return i;
  }
  return -1;
}

// ── Raw 데이터 읽기 (헤더 기반) ─────────────────────────────
function readRawData(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var idx = {
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

  Logger.log('Headers: ' + JSON.stringify(headers));
  Logger.log('Index map: ' + JSON.stringify(idx));

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var result = [];
  var skipped = 0;

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    var dateStr = normalizeDate(idx.date >= 0 ? r[idx.date] : null);
    if (!dateStr) { skipped++; continue; }

    var childAsin = idx.child >= 0 ? String(r[idx.child]).trim() : '';
    if (!childAsin) { skipped++; continue; }

    result.push({
      month:          cellStr(r, idx.month),
      week:           cellStr(r, idx.week),
      weekCode:       cellStr(r, idx.weekCode),
      date:           dateStr,
      parentAsin:     cellStr(r, idx.parent),
      childAsin:      childAsin,
      sessions:       cellNum(r, idx.sessions),
      sessionsB2B:    cellNum(r, idx.sessionsB2B),
      pageViews:      cellNum(r, idx.pageViews),
      pageViewsB2B:   cellNum(r, idx.pageViewsB2B),
      buyBox:         cellNum(r, idx.buyBox),
      units:          cellNum(r, idx.units),
      unitsB2B:       cellNum(r, idx.unitsB2B),
      cvr:            cellNum(r, idx.cvr) * 100,
      sales:          cellNum(r, idx.sales),
      salesB2B:       cellNum(r, idx.salesB2B),
      orderItems:     cellNum(r, idx.orderItems)
    });
  }

  Logger.log('Parsed ' + result.length + ', skipped ' + skipped);
  return result;
}

function cellStr(row, i) {
  if (i < 0 || i >= row.length) return '';
  return String(row[i]).trim();
}

function cellNum(row, i) {
  if (i < 0 || i >= row.length) return 0;
  return toNum(row[i]);
}

// ── 일별 데이터 시트 찾기 ─────────────────────────────────
function findRawDataSheet(ss) {
  var sheets = ss.getSheets();
  var best = null;

  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    if (s.getLastRow() < 2) continue;
    var lastCol = Math.min(s.getLastColumn(), 30);
    var header = s.getRange(1, 1, 1, lastCol).getValues()[0];

    var hasDate = false, hasSessions = false, hasChild = false;
    for (var j = 0; j < header.length; j++) {
      var h = String(header[j]).toLowerCase();
      if (h === 'date') hasDate = true;
      if (h.indexOf('sessions - total') > -1) hasSessions = true;
      if (h.indexOf('child asin') > -1) hasChild = true;
    }
    if (hasDate && hasSessions && hasChild) {
      if (!best || s.getLastRow() > best.getLastRow()) best = s;
    }
  }

  return best || sheets[0];
}

// ── 제품 정보 (ASIN-INFO 탭) ───────────────────────────────
function buildProductMap(ss) {
  var map = {};
  var knownAsins = [];
  var sheets = ss.getSheets();

  var infoSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().replace(/\s/g, '').toUpperCase();
    if (name === 'ASIN-INFO' || name === 'ASININFO' || name === 'PRODUCT' || name === '제품') {
      infoSheet = sheets[i];
      break;
    }
  }

  if (!infoSheet) {
    for (var i = 0; i < sheets.length; i++) {
      var s = sheets[i];
      if (s.getLastRow() < 2) continue;
      var h = s.getRange(1, 1, 1, 3).getValues()[0];
      if (String(h[0]).toUpperCase() === 'ASIN' && String(h[1]).indexOf('상품') > -1) {
        infoSheet = s;
        break;
      }
    }
  }

  if (infoSheet && infoSheet.getLastRow() > 1) {
    var data = infoSheet.getRange(2, 1, infoSheet.getLastRow() - 1, 3).getValues();
    for (var j = 0; j < data.length; j++) {
      var asin  = String(data[j][0]).trim();
      var pName = String(data[j][1]).trim();
      if (asin.indexOf('B0') === 0 && pName) {
        map[asin] = pName;
        knownAsins.push(asin);
      }
    }
  }

  if (knownAsins.length === 0) {
    for (var k in PRODUCT_NAMES) {
      map[k] = PRODUCT_NAMES[k];
      knownAsins.push(k);
    }
  }

  return { map: map, knownAsins: knownAsins };
}

// ── 월 목표 매출 시트 읽기 ────────────────────────────────
function readMonthlyGoals(ss) {
  var goals = {}; // { "2026-05": 50000, ..., "5": 50000, ... }
  var sheets = ss.getSheets();
  var goalSheet = null;

  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().replace(/\s/g, '');
    if (name.indexOf('목표') > -1 || name.toLowerCase().indexOf('goal') > -1 || name.toLowerCase().indexOf('target') > -1) {
      goalSheet = sheets[i];
      break;
    }
  }
  if (!goalSheet || goalSheet.getLastRow() < 1) {
    Logger.log('Goal sheet: not found');
    return goals;
  }

  Logger.log('Goal sheet: ' + goalSheet.getName());
  var lastRow = goalSheet.getLastRow();
  var lastCol = goalSheet.getLastColumn();
  var values  = goalSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // ─── 모든 셀을 스캔하여 (월 텍스트/숫자, 인접 숫자) 페어 찾기 ───
  // 1) 가로 스캔: 같은 행에서 "N월"/"YYYY-MM" + 숫자 페어
  // 2) 세로 스캔: 같은 열의 다음 행이 숫자
  // 3) 헤더 행 + 값 행 (가로)
  // 4) 첫 열 = 월, 둘째 열 = 값 (세로)

  // 케이스 1: 첫 행(또는 두 번째 행)이 헤더 — "1월,2월,...,12월"
  for (var rowH = 0; rowH < Math.min(3, values.length); rowH++) {
    var monthCols = [];
    for (var c = 0; c < values[rowH].length; c++) {
      var key = parseMonthKey_(values[rowH][c]);
      if (key) monthCols.push({ col: c, key: key });
    }
    if (monthCols.length >= 3 && rowH + 1 < values.length) {
      // 다음 행이 숫자값 행으로 추정
      var found = 0;
      for (var k = 0; k < monthCols.length; k++) {
        var v = toNum(values[rowH + 1][monthCols[k].col]);
        if (v) { goals[monthCols[k].key] = v; found++; }
      }
      if (found >= 3) {
        Logger.log('Goals (horizontal): ' + JSON.stringify(goals));
        return goals;
      }
    }
  }

  // 케이스 2: 세로형 — 어느 두 인접 열에서 (월키, 숫자) 페어 찾기
  // 가장 많이 매칭되는 (월컬럼, 값컬럼) 조합을 찾음
  var bestPair = null, bestCount = 0;
  for (var mc = 0; mc < lastCol; mc++) {
    for (var vc = 0; vc < lastCol; vc++) {
      if (mc === vc) continue;
      var tmpGoals = {};
      var cnt = 0;
      for (var r = 0; r < values.length; r++) {
        var key2 = parseMonthKey_(values[r][mc]);
        var val  = toNum(values[r][vc]);
        if (key2 && val) { tmpGoals[key2] = val; cnt++; }
      }
      if (cnt > bestCount) {
        bestCount = cnt;
        bestPair  = tmpGoals;
      }
    }
  }
  if (bestPair && bestCount >= 1) {
    Logger.log('Goals (vertical, ' + bestCount + ' pairs): ' + JSON.stringify(bestPair));
    return bestPair;
  }

  Logger.log('Goal sheet found but no parseable pairs. Values: ' + JSON.stringify(values));
  return goals;
}

function parseMonthKey_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM');
  }
  if (typeof v === 'number') {
    // 1~12면 그냥 월
    if (v >= 1 && v <= 12 && v === Math.floor(v)) return String(v);
    // Excel serial
    var d = new Date((v - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'UTC', 'yyyy-MM');
  }
  var s = String(v).trim();
  if (!s) return '';
  // "2026-05", "2026/05", "2026.05", "2026년 05월"
  var m = s.match(/(\d{4})[년\-\/\.\s]+(\d{1,2})/);
  if (m) return m[1] + '-' + ('0'+m[2]).slice(-2);
  // "5월" 또는 "5" → 월만 (연도 없음)
  m = s.match(/^(\d{1,2})\s*월?$/);
  if (m) {
    var n = parseInt(m[1], 10);
    if (n >= 1 && n <= 12) return String(n);
  }
  return '';
}

// ── 날짜 정규화 ────────────────────────────────────────────
function normalizeDate(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  if (typeof v === 'number') {
    var d = new Date((v - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0'+m[1]).slice(-2) + '-' + ('0'+m[2]).slice(-2);
  var p = new Date(s);
  if (!isNaN(p.getTime())) return Utilities.formatDate(p, 'Asia/Seoul', 'yyyy-MM-dd');
  return '';
}

function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  return parseFloat(String(v).replace(/[$,%\s]/g, '')) || 0;
}
