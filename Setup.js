/***** 初期セットアップ（初回のみ手動実行） *****/

/**
 * 銘柄リストシートを作成してヘッダーを設定する
 * GASエディタで1回だけ手動実行してください
 */
function 初期セットアップ() {
  const ss = getSpreadsheet_();

  // シート作成
  let sh = ss.getSheetByName(FETCH_CONF.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(FETCH_CONF.SHEET_NAME);
  }

  // ヘッダー行
  sh.getRange(1, 1, 1, 13).setValues([[
    'コード', '銘柄名', '株価', '配当利回(%)', '理想利回(%)',
    '判定', '取得対象', '最終更新', '取得元', '保有数量', '取得単価', '前日終値', '配当月'
  ]]);

  // 判定列（F）に数式を設定（3行目以降のサンプル）
  // 実際の銘柄を入力したあとに F2 へ =IF(AND(D2<>"",E2<>"",D2>E2),"買い","様子見") を入力してください

  // 書式設定
  sh.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#e8f0fe');
  sh.setColumnWidth(1, 80);   // コード
  sh.setColumnWidth(2, 160);  // 銘柄名
  sh.setColumnWidth(3, 80);   // 株価
  sh.setColumnWidth(4, 110);  // 配当利回
  sh.setColumnWidth(5, 110);  // 理想利回
  sh.setColumnWidth(6, 80);   // 判定
  sh.setColumnWidth(7, 90);   // 取得対象
  sh.setColumnWidth(8, 140);  // 最終更新
  sh.setColumnWidth(9, 80);   // 取得元
  sh.setColumnWidth(10, 90);   // 保有数量
  sh.setColumnWidth(11, 90);   // 取得単価
  sh.setColumnWidth(12, 90);   // 前日終値
  sh.setColumnWidth(13, 130);  // 配当月（例: 3,9）
  sh.setFrozenRows(1);

  // ScriptPropertiesにスプレッドシートIDを登録
  PropertiesService.getScriptProperties()
    .setProperty('SPREADSHEET_ID', ss.getId());

  Logger.log('初期セットアップ完了。「銘柄リスト」シートに銘柄を入力してください。');
  Logger.log('次に「トリガー登録」を実行してください。');
}


/**
 * 15分ごとの自動更新トリガーを登録する
 * 初期セットアップ後に1回だけ手動実行してください
 */
function トリガー登録() {
  // 既存のupdateAllStocksトリガーを削除（重複防止）
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'updateAllStocks')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('updateAllStocks')
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log('トリガー登録完了: 15分ごとに自動更新します。');
}


/**
 * すべてのトリガーを削除する（リセット用）
 */
function トリガー全削除() {
  ScriptApp.getProjectTriggers()
    .forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('全トリガーを削除しました。');
}
