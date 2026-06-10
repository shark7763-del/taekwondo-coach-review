/**********************************************************************
 * LINE 自動推播 — 教練存檔後自動把復盤回報丟進道館 LINE 群
 *
 * 用 LINE Messaging API（不是已停用的 LINE Notify）。
 * 設定都存在「指令碼屬性」，不用改程式碼：
 *   LINE_TOKEN  = Channel access token（長期）
 *   LINE_TARGET = 要推播的目標 ID（群組 groupId / 使用者 userId）
 *   LINE_AUTO   = 'on'(預設) 或 'off'（暫停自動推播）
 *
 * 取得步驟見 ../LINE推播說明.md
 **********************************************************************/

/* ---- 讀設定 ---- */
function getLineConfig_() {
  var p = PropertiesService.getScriptProperties();
  return {
    token: p.getProperty('LINE_TOKEN') || '',
    target: p.getProperty('LINE_TARGET') || '',
    auto: (p.getProperty('LINE_AUTO') || 'on')
  };
}

/* ---- 一次性設定（在編輯器選這個函式按執行，填好你的值） ---- */
function setupLine() {
  var TOKEN  = '貼上你的 Channel access token';
  var TARGET = '貼上群組 groupId 或你的 userId';
  PropertiesService.getScriptProperties().setProperties({
    LINE_TOKEN: TOKEN,
    LINE_TARGET: TARGET,
    LINE_AUTO: 'on'
  });
  Logger.log('LINE 設定完成：target=' + TARGET);
}
function pauseLinePush()  { PropertiesService.getScriptProperties().setProperty('LINE_AUTO', 'off'); }
function resumeLinePush() { PropertiesService.getScriptProperties().setProperty('LINE_AUTO', 'on'); }

/* ---- 推播主函式（saveReviewSrv 會呼叫） ---- */
function pushToLine_(text, force) {
  var cfg = getLineConfig_();
  if (!cfg.token || !cfg.target) return;
  if (cfg.auto === 'off' && !force) return; // 沒設定就靜默略過
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + cfg.token },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      to: cfg.target,
      messages: [{ type: 'text', text: text }]
    })
  });
  if (res.getResponseCode() >= 300) {
    console.error('LINE push 失敗 ' + res.getResponseCode() + '：' + res.getContentText());
  }
}

/* ---- 測試推播：在編輯器執行這個，確認群組會收到訊息 ---- */
function testLinePush() {
  pushToLine_('🥋 教練復盤系統測試訊息：能看到這則就代表 LINE 推播設定成功！');
}

/**********************************************************************
 * LINE Webhook — 用來「抓出群組 ID」
 * 把 GAS 部署網址(.../exec) 填到 LINE 後台的 Webhook URL，
 * 然後在你的道館群組裡隨便打一句話，群組 ID 會：
 *   1) 寫進試算表新分頁「LineLog」
 *   2) 機器人回一則訊息把 ID 貼給你
 * 拿到 groupId 後填進 setupLine() 的 TARGET，即完成。設定好後可關閉 webhook。
 **********************************************************************/
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var events = body.events || [];
    var sh = sheet_('LineLog', ['time', 'type', 'sourceType', 'sourceId', 'text']);
    var cfg = getLineConfig_();
    events.forEach(function (ev) {
      var src = ev.source || {};
      var srcId = src.groupId || src.roomId || src.userId || '';
      var msg = (ev.message && ev.message.text) || '';
      sh.appendRow([new Date(), ev.type, src.type || '', srcId, msg]);
      // 回覆把 ID 貼回去（需要 token 與 replyToken）
      if (cfg.token && ev.replyToken && ev.type === 'message') {
        UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
          method: 'post', contentType: 'application/json',
          headers: { 'Authorization': 'Bearer ' + cfg.token }, muteHttpExceptions: true,
          payload: JSON.stringify({
            replyToken: ev.replyToken,
            messages: [{ type: 'text', text: '這個對話的 ID：\n' + srcId + '\n\n把它填進 setupLine() 的 TARGET 就完成設定。' }]
          })
        });
      }
    });
  } catch (err) {
    console.error('doPost 錯誤：' + err);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
