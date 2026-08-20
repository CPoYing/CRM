# 讓 Streamlit Cloud 的資料不會不見（接 Google Sheets）

Streamlit Community Cloud **沒有永久磁碟**，容器每次重建都會從 GitHub 重拉一份，
所以寫在硬碟上的 `data/tasks.json` 一定會消失。以下情況都會重建：

- App 閒置睡眠後被喚醒
- 你 push 新 code 觸發重新部署
- Streamlit 那邊維護 / 換機器
- 你自己在後台按 Reboot

設定完以下步驟後，資料就改存在 Google 試算表，怎麼重開都在。
**設定前後你都不用改任何程式碼** —— App 會自己判斷：有 secrets 就用 Google Sheets，沒有就用本機檔案。

---

## 步驟 1：建立試算表

1. 開新分頁輸入 `sheets.new` 建立空白試算表
2. 命名成好認的名字，例如「專案管理開發時程」
3. **複製網址**（整段都要，例如 `https://docs.google.com/spreadsheets/d/1AbC...xyz/edit`）

裡面不用先建工作表，App 第一次啟動會自己建「任務」和「設定」兩個分頁並填入資料。

## 步驟 2：建立 Google Cloud 服務帳號

服務帳號 = 一組給程式用的機器帳號，這樣 App 才有權限讀寫你的試算表。

1. 到 [console.cloud.google.com](https://console.cloud.google.com/)（用你的 Google 帳號登入）
2. 左上角專案選單 → **新增專案** → 取名例如 `project-tracker` → 建立
3. 確認右上角已切換到剛建的專案
4. 搜尋列輸入 **Google Sheets API** → 進去按 **啟用**
5. 搜尋列輸入 **Google Drive API** → 進去按 **啟用**（兩個都要）
6. 左側選單 → **API 和服務** → **憑證** → 上方 **建立憑證** → **服務帳戶**
7. 服務帳戶名稱隨便取（例如 `tracker-bot`）→ 建立並繼續 → 角色可略過 → 完成
8. 回到憑證頁，點剛建立的服務帳戶 → 上方 **金鑰** 分頁 → **新增金鑰** → **建立新的金鑰** → 選 **JSON** → 建立
9. 瀏覽器會下載一個 `.json` 檔，**這個檔案就是密碼，不要傳給別人、不要放進 GitHub**

## 步驟 3：把試算表分享給服務帳號

1. 用記事本打開剛下載的 JSON，找到 `"client_email"`，長得像
   `tracker-bot@project-tracker-123456.iam.gserviceaccount.com`
2. 回到步驟 1 的試算表 → 右上角 **共用**
3. 貼上那個 email，權限選 **編輯者** → 傳送
   （會跳「這不是 Google 帳戶」之類的提醒，照樣送出即可）

**漏掉這步是最常見的失敗原因**，App 會出現 `SpreadsheetNotFound` 或 403。

## 步驟 4：填 Streamlit Cloud 的 Secrets

1. 到 [share.streamlit.io](https://share.streamlit.io) → 你的 App → 右下 **⋮** → **Settings** → **Secrets**
2. 貼上下面的內容，把 `<>` 的部分換成你自己的（值都從那個 JSON 檔複製）：

```toml
[gsheets]
spreadsheet = "<步驟 1 複製的試算表網址>"

[gcp_service_account]
type = "service_account"
project_id = "<JSON 裡的 project_id>"
private_key_id = "<JSON 裡的 private_key_id>"
private_key = "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBg...(很長)...\n-----END PRIVATE KEY-----\n"
client_email = "<JSON 裡的 client_email>"
client_id = "<JSON 裡的 client_id>"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "<JSON 裡的 client_x509_cert_url>"
universe_domain = "googleapis.com"
```

3. **Save** → App 會自動重啟

### ⚠️ `private_key` 是最容易踩雷的一行

- 必須是**單獨一行**，用雙引號包起來
- 裡面的換行要保持成 **`\n` 兩個字元**，不要真的按 Enter 斷行
- 直接從 JSON 檔複製那一整串（JSON 裡本來就已經是 `\n` 的形式）最保險
- 開頭結尾的 `-----BEGIN PRIVATE KEY-----` / `-----END PRIVATE KEY-----` 不能少

## 步驟 5：確認

打開 App，左側欄最下面應該顯示：

```
儲存方式：Google Sheets
資料存在雲端試算表，重新部署或睡眠都不會不見
```

然後隨便改一格，回試算表看 —— 資料應該同步出現。
再按 Streamlit 後台的 **Reboot**，重開後資料還在就成功了。

---

## 本機也想連同一份試算表（選用）

在 `project_tracker_app/` 底下建立 `.streamlit/secrets.toml`，內容跟步驟 4 一樣。
這個路徑已經在 `.gitignore` 裡，不會被推上 GitHub。

不建立的話，本機就繼續用 `data/tasks.json`，跟雲端各自獨立、互不影響。

---

## 常見問題

| 症狀 | 原因 |
| --- | --- |
| `SpreadsheetNotFound` | 步驟 3 沒做，試算表沒分享給服務帳號的 email |
| `403 Google Drive API has not been used` | 步驟 2-5 的 Drive API 沒啟用 |
| `Incorrect padding` / `Could not deserialize key` | `private_key` 換行壞掉，見上面的雷區說明 |
| 左側欄顯示「本機檔案」 | Secrets 沒存好，或 `[gsheets]`、`[gcp_service_account]` 少了一個 |
| 存檔偶爾失敗 | Google 每分鐘寫入次數上限，程式會自動重試；連續狂改很多格才可能碰到 |

## 這樣算多人協作了嗎？

**還沒有。** 資料是共用的沒錯，但：

- 沒有登入機制，拿到 App 網址的人都能改
- 兩個人同時開，各自的畫面是修改前的副本，**誰後存檔誰覆蓋**
- 別人改過之後，你要按左側「重新載入雲端資料」才看得到

比較穩的用法是：**一個人在 App 裡維護，其他人看**；
或是**大家直接在 Google 試算表裡編輯**（試算表原生就支援多人即時協作、有版本紀錄），
App 純粹當統計和甘特圖的檢視介面。
