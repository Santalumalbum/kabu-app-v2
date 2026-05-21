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

    const rows = sh.getRange(2, 1, lastRow - 1, 12).getDisplayValues();
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
        owned:     String(r[9]).trim(),
        cost:      String(r[10]).trim(),
        prevClose: String(r[11]).trim(), // 前日終値
      }));

    // ポートフォリオ集計
    let totalValue = 0, totalDividend = 0;
    items.forEach(item => {
      const owned = parseFloat(item.owned);
      const price = parseFloat(item.price.replace(/,/g, ''));
      const divY  = parseFloat(item.divYield);
      if (owned > 0 && !isNaN(price)) totalValue    += owned * price;
      if (owned > 0 && !isNaN(price) && !isNaN(divY))
        totalDividend += owned * price * (divY / 100);
    });

    return {
      ok: true, items: items, updatedAt: updatedAt,
      portfolio: {
        totalValue:    Math.round(totalValue),
        totalDividend: Math.round(totalDividend),
      },
    };

  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}

/**
 * 手動更新ボタンから呼ばれる（市場時間外でも強制実行）
 */
function runManualUpdate() {
  try {
    updateAllStocksForced_();
    return getDisplayData();
  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}
