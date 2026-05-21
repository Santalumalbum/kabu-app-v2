/***** 株価・配当利回り自動取得 *****/

const FETCH_CONF = {
  SHEET_NAME:      '銘柄リスト',
  START_ROW:       2,
  COL_CODE:        1,  // A: 銘柄コード
  COL_NAME:        2,  // B: 銘柄名
  COL_PRICE:       3,  // C: 株価
  COL_YIELD:       4,  // D: 配当利回(%)
  COL_IDEAL:       5,  // E: 理想利回(%)
  COL_CHECK:       7,  // G: 取得対象チェック
  COL_UPDATED:     8,  // H: 最終更新日時
  COL_SOURCE:      9,  // I: 取得元
  PRICE_CACHE_SEC: 180,   // 株価キャッシュ 3分
  YIELD_CACHE_SEC: 1800,  // 利回りキャッシュ 30分
  BATCH_SIZE:      10,
  BATCH_SLEEP_MS:  1000,
};


/**
 * 現在が東証の取引時間内かどうか判定
 * 09:00-11:30 / 12:30-15:30（土日除外）
 */
function isMarketHours_() {
  const now = new Date();
  const dow = now.getDay();
  if (dow === 0 || dow === 6) return false;

  const hhmm = now.getHours() * 100 + now.getMinutes();
  return (hhmm >= 900 && hhmm <= 1130) || (hhmm >= 1230 && hhmm <= 1530);
}

/**
 * チェック済み銘柄の株価・利回りを一括更新
 * トリガーから呼び出す（市場時間外はスキップ）
 */
function updateAllStocks() {
  if (!isMarketHours_()) {
    Logger.log('市場時間外のためスキップ');
    return;
  }
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(FETCH_CONF.SHEET_NAME);
  if (!sh) {
    Logger.log('「銘柄リスト」シートが見つかりません。');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < FETCH_CONF.START_ROW) return;

  const numRows = lastRow - FETCH_CONF.START_ROW + 1;
  const range = sh.getRange(FETCH_CONF.START_ROW, 1, numRows, FETCH_CONF.COL_SOURCE);
  const values = range.getValues();

  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const code    = String(row[FETCH_CONF.COL_CODE - 1]  || '').trim();
    const checked = row[FETCH_CONF.COL_CHECK - 1];

    if (!code || !isChecked_(checked)) continue;

    const result = fetchStockData_(code);
    const now = new Date();

    // 取得成功した値だけ上書き（失敗時は前回値を保持）
    if (result.price !== null) {
      sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_PRICE).setValue(result.price);
    }
    if (result.divYield !== null) {
      sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_YIELD).setValue(result.divYield);
    }

    const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
    const sourceStr = result.source + (result.price === null ? '(株価取得失敗)' : '')
                                    + (result.divYield === null ? '(利回取得失敗)' : '');
    sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_UPDATED).setValue(timeStr);
    sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_SOURCE).setValue(sourceStr);
    if (result.prevClose !== null) {
      sh.getRange(FETCH_CONF.START_ROW + i, PORTFOLIO_COLS.COL_PREV_CLOSE).setValue(result.prevClose);
    }

    count++;

    if (count % FETCH_CONF.BATCH_SIZE === 0) {
      SpreadsheetApp.flush();
      Utilities.sleep(FETCH_CONF.BATCH_SLEEP_MS);
    }
  }

  // H1セルに最終更新日時を記録
  sh.getRange('H1').setValue(
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm 更新')
  );

  SpreadsheetApp.flush();
  Logger.log('更新完了: ' + count + '銘柄');
}


/**
 * 1銘柄の株価・利回りを取得
 */
function fetchStockData_(code) {
  const result = { price: null, divYield: null, source: '' };
  const padded = code.padStart(4, '0');

  // --- 株価取得（Yahoo Finance JSON → Google Finance）---
  const price = fetchPriceYahooJson_(padded) || fetchPriceGoogleFinance_(padded);
  result.price = price;
  result.prevClose = fetchPrevClose_(padded);

  // --- 利回り取得（Yahoo Finance JP → minkabu）---
  const yieldData = fetchYieldYahoo_(padded) || fetchYieldMinkabu_(padded);
  if (yieldData !== null) {
    result.divYield = yieldData;
    result.source = '取得済';
  } else {
    result.source = 'FAIL';
  }

  return result;
}


/***** 株価取得 *****/

function fetchPrevClose_(code) {
  try {
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T'
              + '?interval=1d&range=5d';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const json = JSON.parse(res.getContentText());
    const closes = json.chart && json.chart.result && json.chart.result[0]
      ? json.chart.result[0].indicators.quote[0].close : null;
    if (!closes || closes.length < 2) return null;
    // 直近の有効な前日終値を返す
    const valid = closes.filter(v => v !== null);
    return valid.length >= 2 ? valid[valid.length - 2] : null;
  } catch (e) {
    return null;
  }
}

function fetchPriceYahooJson_(code) {
  const cacheKey = 'p_yj_' + code;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return parseFloat(cached);

  try {
    const url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + code + '.T'
              + '?interval=1d&range=1d';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    const json = JSON.parse(res.getContentText());
    const meta = json.chart && json.chart.result && json.chart.result[0]
               ? json.chart.result[0].meta : null;
    const price = meta ? (meta.regularMarketPrice || meta.previousClose) : null;
    if (!price) return null;

    cache.put(cacheKey, String(price), FETCH_CONF.PRICE_CACHE_SEC);
    return price;
  } catch (e) {
    Logger.log('YahooJSON price error ' + code + ': ' + e.message);
    return null;
  }
}

function fetchPriceGoogleFinance_(code) {
  const cacheKey = 'p_gf_' + code;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return parseFloat(cached);

  try {
    const url = 'https://www.google.com/finance/quote/' + code + ':TYO';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    const html = res.getContentText();
    const m = html.match(/data-last-price="([\d.]+)"/);
    if (!m) return null;

    const price = parseFloat(m[1]);
    cache.put(cacheKey, String(price), FETCH_CONF.PRICE_CACHE_SEC);
    return price;
  } catch (e) {
    Logger.log('GF price error ' + code + ': ' + e.message);
    return null;
  }
}


/***** 利回り取得 *****/

function fetchYieldYahoo_(code) {
  const cacheKey = 'y_yj_' + code;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached === 'null') return null;
  if (cached) return parseFloat(cached);

  try {
    const url = 'https://finance.yahoo.co.jp/quote/' + code + '/dividend';
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    const html = res.getContentText();

    // 「配当利回り」を探す
    const patterns = [
      /配当利回り[^<]*<\/[^>]+>\s*<[^>]+>([\d.]+)\s*%/,
      /予想配当利回り[^<]*<\/[^>]+>\s*<[^>]+>([\d.]+)\s*%/,
      /"dividendYield"[^>]*>([\d.]+)/,
    ];
    for (const pat of patterns) {
      const m = html.match(pat);
      if (m) {
        const val = parseFloat(m[1]);
        cache.put(cacheKey, String(val), FETCH_CONF.YIELD_CACHE_SEC);
        return val;
      }
    }

    cache.put(cacheKey, 'null', FETCH_CONF.YIELD_CACHE_SEC);
    return null;
  } catch (e) {
    Logger.log('Yahoo yield error ' + code + ': ' + e.message);
    return null;
  }
}

function fetchYieldMinkabu_(code) {
  const cacheKey = 'y_mk_' + code;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached === 'null') return null;
  if (cached) return parseFloat(cached);

  try {
    const url = 'https://minkabu.jp/stock/' + code;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;

    const html = res.getContentText();
    const patterns = [
      /予想配当利回り[^<]*<\/[^>]+>\s*<[^>]+>([\d.]+)/,
      /配当利回り[^<]*<\/[^>]+>\s*<[^>]+>([\d.]+)/,
      /"dividendYield"[^:]*:\s*"?([\d.]+)/,
    ];
    for (const pat of patterns) {
      const m = html.match(pat);
      if (m) {
        const val = parseFloat(m[1]);
        cache.put(cacheKey, String(val), FETCH_CONF.YIELD_CACHE_SEC);
        return val;
      }
    }

    cache.put(cacheKey, 'null', FETCH_CONF.YIELD_CACHE_SEC);
    return null;
  } catch (e) {
    Logger.log('minkabu yield error ' + code + ': ' + e.message);
    return null;
  }
}


/***** ユーティリティ *****/

/**
 * 市場時間チェックなしで強制更新（手動ボタン用）
 */
function updateAllStocksForced_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(FETCH_CONF.SHEET_NAME);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < FETCH_CONF.START_ROW) return;

  const numRows = lastRow - FETCH_CONF.START_ROW + 1;
  const values = sh.getRange(FETCH_CONF.START_ROW, 1, numRows, FETCH_CONF.COL_SOURCE).getValues();
  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const code    = String(values[i][FETCH_CONF.COL_CODE - 1] || '').trim();
    const checked = values[i][FETCH_CONF.COL_CHECK - 1];
    if (!code || !isChecked_(checked)) continue;

    const result = fetchStockData_(code);
    const now = new Date();
    if (result.price    !== null) sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_PRICE).setValue(result.price);
    if (result.divYield !== null) sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_YIELD).setValue(result.divYield);
    sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_UPDATED)
      .setValue(Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'));
    sh.getRange(FETCH_CONF.START_ROW + i, FETCH_CONF.COL_SOURCE).setValue(result.source);
    count++;
    if (count % FETCH_CONF.BATCH_SIZE === 0) { SpreadsheetApp.flush(); Utilities.sleep(FETCH_CONF.BATCH_SLEEP_MS); }
  }

  sh.getRange('H1').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm 更新'));
  SpreadsheetApp.flush();
}

function isChecked_(val) {
  if (val === true || val === 'TRUE') return true;
  if (typeof val === 'number' && val !== 0) return true;
  if (typeof val === 'string' && val.trim() !== '' && val !== 'FALSE') return true;
  return false;
}
