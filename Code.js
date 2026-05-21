/***** Webアプリ入口 *****/

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('株アプリv2')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getDisplayData() {
  try {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName('銘柄リスト');
    if (!sh) return { ok: false, message: '「銘柄リスト」シートが見つかりません。' };

    const lastRow = sh.getLastRow();
    const updatedAt = sh.getRange('H1').getDisplayValue();

    if (lastRow < 2) {
      return { ok: true, items: [], updatedAt: updatedAt };
    }

    const rows = sh.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    // A:コード B:銘柄名 C:株価 D:配当利回 E:理想利回 F:判定 G:メモ

    const items = rows
      .filter(r => r[0] || r[1])
      .map(r => ({
        code:     r[0],
        name:     r[1],
        price:    r[2],
        divYield: r[3],
        ideal:    r[4],
        judge:    r[5],
        memo:     r[6],
      }));

    return { ok: true, items: items, updatedAt: updatedAt };

  } catch (err) {
    Logger.log(err.stack || err);
    return { ok: false, message: err.message };
  }
}
