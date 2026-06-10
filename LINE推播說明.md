# LINE 自動推播設定教學

目標：教練在系統按「儲存」後，後端自動把該堂的 LINE 回報文字推進道館 LINE 群組，**教練不用再手動複製貼上**。

> 用的是 **LINE Messaging API**（LINE Notify 已於 2025 年停用，不要再找它）。
> 程式都在 `apps-script/LineBot.gs`，設定值存在「指令碼屬性」，不用改程式碼。

---

## 一、建立 LINE 官方帳號 + Messaging API
1. 到 [LINE Developers](https://developers.line.biz/) 用 LINE 帳號登入。
2. 建立一個 **Provider**（例如「雄麒道館」）。
3. 在該 Provider 下 **Create a new channel → Messaging API**，填道館資訊。
4. 進入頻道 → **Messaging API** 分頁：
   - 最下方 **Channel access token (long-lived)** 按 **Issue** 產生並複製 →（這是 `LINE_TOKEN`）
   - **Auto-reply / 加入問候** 可關掉，避免洗版。

## 二、把官方帳號加進道館群組
1. 用 **Messaging API** 分頁上方的 QR code 加這個官方帳號為好友。
2. 把這個官方帳號**邀請進你的道館教練 LINE 群組**。
   （Messaging API 推播到群組，前提是 bot 在群裡。）

## 三、抓出群組 ID（用 Webhook 一次性取得）
群組 ID 不會直接顯示，要用 webhook 抓：
1. GAS 編輯器執行 `setupLine()` 前，**先只填 TOKEN**：
   - 打開 `LineBot.gs`，把 `setupLine()` 裡的 `TOKEN` 換成你的 token，`TARGET` 先隨便留著，存檔，執行 `setupLine()` 授權。
2. 到 LINE Developers → Messaging API → **Webhook URL** 填你的 GAS 部署網址（`.../exec`），按 **Verify**，並把 **Use webhook** 打開。
3. 在道館群組裡**隨便打一句話**。
4. 機器人會回一則訊息把「這個對話的 ID」貼出來；同時試算表也會多一個 **LineLog** 分頁記著 `sourceId`。複製那個 `Cxxxxx...`（群組 ID 以 `C` 開頭）。

## 四、完成設定
1. 回 GAS 編輯器，把 `setupLine()` 裡的 `TARGET` 換成剛剛的群組 ID。
2. 執行 `setupLine()`。
3. 執行 `testLinePush()` → 群組若收到「測試訊息」就成功了。
4. 設定完成後 LINE 後台的 **Use webhook 可關掉**（抓 ID 才需要）。

---

## 日常使用
- 教練在系統按「💾 儲存並產生 LINE 回報」→ 後端自動推播到群組。
- 只有「**新存的一筆**」會推；後台覆寫同一筆、載入示範資料不會洗版。
- 想暫停／恢復自動推播：在 GAS 執行 `pauseLinePush()` / `resumeLinePush()`，或把指令碼屬性 `LINE_AUTO` 設為 `off` / `on`。

## 疑難排解
| 狀況 | 處理 |
|---|---|
| 群組收不到 | 確認官方帳號**在群組裡**、`LINE_TARGET` 是 `C` 開頭的群組 ID、`testLinePush()` 有沒有報錯 |
| `testLinePush` 報 401 | token 錯或過期，重新 Issue Channel access token |
| `testLinePush` 報 400 | `LINE_TARGET` 格式不對（要 groupId/userId，不是顯示名稱） |
| 抓不到群組 ID | Webhook URL 要用 `.../exec`、Use webhook 要開、官方帳號要在群裡且訊息是發在那個群 |
| 免費額度 | Messaging API 免費方案每月有推播則數上限，一般道館每天數則遠低於上限 |
