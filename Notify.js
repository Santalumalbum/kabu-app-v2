/***** LINE通知 *****/

function sendLineMessage_(message, userId) {
  const token = getLineAccessToken_();
  if (!token) {
    Logger.log('LINE_ACCESS_TOKEN が未設定です。');
    return false;
  }

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token,
    },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: message }],
    }),
    muteHttpExceptions: true,
  };

  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    const code = res.getResponseCode();
    Logger.log('LINE送信: ' + code + ' → ' + userId);
    return code >= 200 && code < 300;
  } catch (err) {
    Logger.log('LINE送信エラー: ' + err.message);
    return false;
  }
}

/**
 * 買い候補をまとめてLINE送信
 */
function notifyBuyCandidates_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('銘柄リスト');
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  // A:コード B:銘柄名 C:株価 D:配当利回 E:理想利回 F:判定

  const header = '銘柄/コード/配当利回/理想利回';
  const lines = [];

  for (const row of data) {
    const [code, name, , divY, ideal, judge] = row;
    if (String(judge).indexOf('買') !== -1 && code && name) {
      lines.push(`${code} ${name} / ${Number(divY).toFixed(2)} / ${Number(ideal).toFixed(2)}`);
    }
  }

  if (lines.length === 0) {
    Logger.log('買い候補なし');
    return;
  }

  const MAX_LEN = 1800;
  const chunks = [];
  let buf = header;

  for (const line of lines) {
    if ((buf + '\n' + line).length > MAX_LEN) {
      chunks.push(buf);
      buf = header + '\n' + line;
    } else {
      buf += '\n' + line;
    }
  }
  if (buf) chunks.push(buf);

  const userIds = getStockUserIds_();
  for (const uid of userIds) {
    for (const msg of chunks) {
      sendLineMessage_(msg, uid);
    }
  }

  Logger.log(`通知完了: ${userIds.length}人 × ${chunks.length}件`);
}
