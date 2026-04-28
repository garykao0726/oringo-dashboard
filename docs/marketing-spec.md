# 林果良品行銷儀錶板 — 規格與現況

## 系統定位

行銷儀錶板（給行銷團隊看），與「營運儀錶板」(`index.html`) 分離。
- 入口：`marketing.html`
- 資料：延用「林果良品流量廣告監控」GAS 專案 + 同名 Google Sheet
- SEO 內容：以分頁 / iframe 形式嵌入既有 `seo.html`

## 系統架構

```
┌──────────────────────────────────────────────┐
│  Google Apps Script「林果良品流量廣告監控」      │
│  - 既有：weeklyAdReport / Gmail POS 掃描       │
│  - 新增：每日自動拉取 API 寫入統一資料表          │
└──────────────────────────────────────────────┘
              ↓
┌──────────────────────────────────────────────┐
│  Google Sheets                                │
│  - 既有：每週輸入 / 門市月報_POS / 門市月報_Shopline │
│  - 新增：行銷儀錶板數據（每日一行）                 │
└──────────────────────────────────────────────┘
              ↓ gviz/CSV
┌──────────────────────────────────────────────┐
│  GitHub Pages                                  │
│  - marketing.html（新）                        │
│  - seo.html（既有，被嵌入）                      │
└──────────────────────────────────────────────┘
```

## 資料來源

| 來源 | API / 方式 | 主要指標 |
|---|---|---|
| GA4 Data API | Property `259936795` | sessions, newUsers, ecommercePurchases, purchaseRevenue, addToCarts, checkouts, sessionSource (IG/FB/LINE) |
| Shopline (open API) | `open.shopline.io/v1` | 訂單、營收、退貨 |
| Meta Ads | `graph.facebook.com/v19.0` | spend, ROAS, CTR |
| Google Ads | `googleads.googleapis.com/v16` | cost, conversions_value |
| Gmail POS | 既有 `checkGmailForPOS` | 各店淨營收 |

## 資料表結構（規劃中，待前端確認後定案）

工作表：`行銷儀錶板數據`，每日一行（不是規格原本的每小時 — 大部分指標日頻夠用，省 API 配額）。

欄位概要：
- 日期、更新時間
- 流量：sessions / newUsers / 林果在哪裡 PV
- 營收：電商 + 4 店 POS / 訂單數 / 客單價 / 退貨率
- 行為：購物車放棄率 / paid CVR / site CVR
- 社群：IG / FB / LINE 各自的 sessions、CVR、轉換金額
- 廣告：Meta spend / ROAS / CTR、Google spend / ROAS / CTR
- 全站 ROAS：(各通路營收) / (Meta + Google 花費)

## 月目標（已內建於 GAS）

```
電商: 125 115 115 100 100 90 90 90 100 115 115 125
中山: 180 170 170 150 140 125 125 125 140 170 170 180
松菸: 160 150 150 140 120 110 110 110 120 140 150 160
台中: 130 120 120 120 110 95 95 95 110 120 120 130
東門: 140 130 130 120 110 100 100 100 120 130 130 140
```
（單位：萬，索引 0=Jan）

## 廣告門檻

- ROAS：目標 4.5 / 警告 3.5 / 警報 2.5
- 整站 CVR：目標 0.35 / 警告 0.28 / 警報 0.20
- Paid CVR：目標 0.59 / 警告 0.45 / 警報 0.30

## 開發階段

**階段 1：前端殼（先做）**
- `marketing.html` 骨架：KPI 卡 / ROAS / 各通路花費 / 流量來源 / SEO iframe 分頁
- 先用 mock data，看到雛型再決定資料結構

**階段 2：後端資料管道**
- GAS 加 `pullDailyMarketingData()`：每日 02:00 拉所有 API 寫入「行銷儀錶板數據」分頁
- 不動既有 `weeklyAdReport`、`checkGmailForPOS`

**階段 3：週報重構**
- `weeklyAdReport` 改成讀「行銷儀錶板數據」分頁，停用「每週輸入」手動數據
- 訊息範本擴充（IG / FB / LINE / 各店進度）

## 安全性

- API token / Webhook URL 全部不進 git（`.gitignore` 已排除 `apps-script.gs` + `marketing-gas/`）
- GAS 雲端 = 真理來源，本機透過 `clasp pull / push` 同步
- 既有 token 已暴露（公開 repo 9 天），採監控策略，異常時再輪換

## 週報訊息範本（階段 3 用）

```
📊 林果良品週報 (W{週次}, {日期區間})

✅ 本週表現
- 全站 ROAS: {roas}（目標 ≥4.5）
- 官網營收: NT${revenue}（週目標 NT${target}，達成 {pct}%）
- 付費流量 CVR: {cvr}%

📈 同比
- 營收 vs 上週: {mom}%
- 營收 vs 去年同期: {yoy}%

💰 廣告投放
- Meta: NT${meta_spend} | ROAS {meta_roas}
- Google: NT${google_spend} | ROAS {google_roas}

🔄 流量來源
- IG: {sessions} 人 | CVR {cvr}%
- FB: {sessions} 人 | CVR {cvr}%
- LINE: {sessions} 人 | CVR {cvr}%

🎯 本月進度（已過 {days}/{total} 天）
- 電商 / 中山 / 松菸 / 台中 / 東門 達成率
```

狀態圖示：✅ ≥90% / ⚠️ 70-90% / 🔴 <70%
