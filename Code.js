/***** Webアプリ入口 *****/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('株ダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * 銘柄リストのデータをWebアプリに返す
 */
function getDisplayData() {
  try {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(FETCH_CONF.SHEET_NAME);
    if (!sh) return { ok: false, message: '「銘柄リスト」シートが見つかりません。まず初期セットアップを実行してください。' };

    const lastRow = sh.getLastRow();
    const updatedAt = sh.getRange('H1').getDisplayValue();

    if (lastRow < 2) return { ok: true, items: [], updatedAt: updatedAt };

    const rows = sh.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
    const items = rows
      .filter(r => String(r[0]).trim() || String(r[1]).trim())
      .map(r => ({
        code:     String(r[0]).trim(),
        name:     String(r[1]).trim(),
        price:    String(r[2]).trim(),
        divYield: String(r[3]).trim(),
        ideal:    String(r[4]).trim(),
        judge:    String(r[5]).trim(),
        source:   String(r[8]).trim(),
      }));

    return { ok: true, items: items, updatedAt: updatedAt };

  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}

/**
 * 手動更新ボタンから呼ばれる
 */
function runManualUpdate() {
  try {
    updateAllStocks();
    return getDisplayData();
  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}
