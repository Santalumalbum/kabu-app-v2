/***** 配当受取記録 *****/

const DR_SHEET = '配当受取記録';
const DR_HEADERS = ['受取日', '銘柄コード', '銘柄名', '受取金額(円)', 'メモ'];

/**
 * 受取記録シートを取得（なければ作成）
 */
function getDividendRecordSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(DR_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DR_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([DR_HEADERS]);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#fce8d5');
    sh.setColumnWidth(1, 110);
    sh.setColumnWidth(2, 80);
    sh.setColumnWidth(3, 160);
    sh.setColumnWidth(4, 110);
    sh.setColumnWidth(5, 200);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * 配当受取を1件保存
 */
function saveDividendRecord(date, code, name, amount, memo) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(8000)) return { ok: false, message: '保存中です。少し待ってください。' };
    try {
      const sh = getDividendRecordSheet_();
      sh.appendRow([date, String(code), String(name), Number(amount), String(memo || '')]);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}

/**
 * 配当受取記録を全件取得して集計
 */
function getDividendRecords() {
  try {
    const sh = getDividendRecordSheet_();
    const last = sh.getLastRow();

    if (last < 2) {
      return { ok: true, records: [], stats: buildEmptyStats_() };
    }

    const rows = sh.getRange(2, 1, last - 1, 5).getValues();
    const records = rows
      .filter(r => r[0] && r[3])
      .map(r => ({
        date:   r[0] instanceof Date
          ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy/MM/dd')
          : String(r[0]),
        code:   String(r[1] || ''),
        name:   String(r[2] || ''),
        amount: Number(r[3]) || 0,
        memo:   String(r[4] || ''),
      }))
      .sort((a, b) => b.date.localeCompare(a.date)); // 新しい順

    const stats = buildStats_(records);
    return { ok: true, records: records, stats: stats };

  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}

/**
 * 統計を集計
 */
function buildStats_(records) {
  const total = records.reduce((s, r) => s + r.amount, 0);

  // 月別集計（直近12ヶ月）
  const monthlyMap = {};
  records.forEach(r => {
    const ym = r.date.slice(0, 7); // "2025-03"
    monthlyMap[ym] = (monthlyMap[ym] || 0) + r.amount;
  });

  // 直近12ヶ月のキーを生成
  const now = new Date();
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    months.push(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM'));
  }
  const monthly = months.map(ym => ({
    ym:     ym,
    label:  ym.slice(5) + '月',
    amount: monthlyMap[ym] || 0,
  }));

  // 銘柄別集計 TOP5
  const byStock = {};
  records.forEach(r => {
    const key = r.code || r.name;
    if (!byStock[key]) byStock[key] = { name: r.name || r.code, amount: 0 };
    byStock[key].amount += r.amount;
  });
  const topStocks = Object.values(byStock)
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 5);

  // 今年の合計
  const thisYear = String(now.getFullYear());
  const thisYearTotal = records
    .filter(r => r.date.startsWith(thisYear))
    .reduce((s, r) => s + r.amount, 0);

  return {
    total:         Math.round(total),
    thisYearTotal: Math.round(thisYearTotal),
    count:         records.length,
    monthly:       monthly,
    topStocks:     topStocks,
  };
}

function buildEmptyStats_() {
  const now = new Date();
  const months = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const ym = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM');
    months.push({ ym: ym, label: ym.slice(5) + '月', amount: 0 });
  }
  return { total: 0, thisYearTotal: 0, count: 0, monthly: months, topStocks: [] };
}
