/***** 設定 *****/

function getSpreadsheetId_() {
  return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
    || '1vH2_9WszgFlEE4LBzcWfnuBb1YWT_aUEBony2ubQnSQ';
}

function getLineAccessToken_() {
  return PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN') || '';
}

function getStockUserIds_() {
  const props = PropertiesService.getScriptProperties();
  const keys = [
    'LINE_USER_ID_1',
    'LINE_USER_ID_2',
    'LINE_USER_ID_3',
    'LINE_USER_ID_4',
    'LINE_USER_ID_5',
  ];
  return keys.map(k => props.getProperty(k)).filter(Boolean);
}

// 保有管理列・前日比列
const PORTFOLIO_COLS = {
  COL_OWNED:     10,  // J: 保有数量
  COL_COST:      11,  // K: 取得単価
  COL_PREV_CLOSE: 12, // L: 前日終値
};

function getSpreadsheet_() {
  return SpreadsheetApp.openById(getSpreadsheetId_());
}
