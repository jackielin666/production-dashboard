---
name: monthly-report
description: 產生「生產鍋數月報」——依 data/latest.xlsx 分析指定月份的業務量（鍋數為主、重量為輔）、分級接單下滑警示、斷單、資料異常、上期追蹤與季節性，輸出 HTML 報告到 reports/ 並推上 GitHub Pages 取得可分享連結；也可在月底前產生「期中預覽」（同期比較）。當使用者說「幫我分析 2026年8月」「幫我出 8 月月報」「分析上個月」「月度分析」「月報」，或「幫我出 9 月期中預覽」「月底開會先看月報」「本月到目前為止」時使用。
---

# 生產鍋數月報

使用者每月初把上月完整的「產品製成率」Excel 覆蓋到 `data/latest.xlsx`，然後說
「幫我分析 YYYY年M月」，你就依本流程產出一份業務導向的月報並回覆連結。

月底廠務會議前，使用者會先匯入最新資料並說「幫我出 M 月期中預覽」，
則走下方「期中預覽」流程（同期比較、標註本月尚未結束）。

**分析取向**：公司是業務導向、接單式生產。**生產鍋數直接反映業務量與營收** ——
某品項鍋數明顯下降 = 出貨量下降 = 接單量下滑 = 營收風險，主力品項尤其重要。
**製成率／品質不在本報告範圍**（由「產品工時及製成率統計系統」負責），不要加回來。
**鍋數為主、重量為輔**：重量可能因製程條件改變而變動，只作參考，**不要**用「每鍋重量變大／變小」
推論批量或業務結論。

**已發布的月報不重新產生**（2026/06~08 維持原樣）；新規則自 2026/09 起適用。

## 步驟

1. **確認資料**：`data/latest.xlsx` 是否已包含目標月份。若使用者剛上傳，先 `git pull`。

2. **產生報告**：
   ```bash
   python3 tools/monthly_report.py --month YYYY-MM
   ```
   - 目標月份不存在 → 會列出資料涵蓋範圍，請使用者先上傳新檔。
   - 目標月份未補齊 → 會擋下並說明筆數。**先回報使用者**，不要逕自加
     `--allow-incomplete`；除非使用者明確說要看未完整的資料。

3. **讀數據寫總評**：讀 `reports/YYYY-MM.json`，寫一段 3~5 段的中文總評，
   存成 `reports/YYYY-MM.narrative.txt`（留在 repo 內，日後可原樣重新產生報告），再執行：
   ```bash
   python3 tools/monthly_report.py --month YYYY-MM --narrative reports/YYYY-MM.narrative.txt
   ```
   總評要回答老闆真正關心的問題，而不是複述數字：
   - **整體業務量方向**：環比／同比，用 `waterfall` 說明變動主要來自哪些品項，
     並用 `seasonal_note` 判斷是季節性還是實質衰退（`seasonal_sample` 樣本少於 3 年時要說「僅供參考」）。
   - **YTD**：用 `ytd` 區塊。有季節性代工時，**主數字用排除後的 `ytd.pots`／`ytd.pct`**，
     並清楚註明排除了哪些品項（`oem_names`、`oem_pots`）以及含代工的數字（`all_pots`／`all_pct`）。
   - **營收風險**：`grade` 為 **A**（主力）的警示必須逐項點名；B 級挑重點；C 級不必提。
     `dormant` 中 `suspect_lost`（連續未生產 ≥3 月）要點名為「疑似流失」。
   - **`verdict`（重量參考）不影響判斷**：不要因為重量持平就把鍋數下降說成「不是接單下滑」。
   - **改版提示**（`hints`）：有提示時可說「可能為改版轉移，請業務確認」，但**不要**自行認定為同一產品。
   - **資料異常**（`anomalies`）：只點出品項與數字，**不解釋原因**，寫明「請製造單位說明」。
     `kg_pending` 為 true 時代表重量尚未填入，不要評論重量。
   - **上期追蹤**（`tracking`）：簡述上期點名事項的進度（恢復／持續／惡化）。
   - **成長面**：哪些品項在補上缺口。
   - **注意品項註記**（`annotated` 欄）：標為「已下市」的品項是計畫性下市，
     **不可**當成接單流失；標為「季節性代工」的品項量體大且間歇，
     **不可**當成本業的成長動能或衰退風險。若報告有 `ex_oem` 欄，
     請以「排除季節性代工後」的數字描述本業表現。
   - **建議追查方向**：例如請業務確認某客戶／品項的後續訂單。

4. **提交並回連結**：
   ```bash
   git add reports/ && git commit && git push -u origin <branch>
   ```
   （依專案慣例：在指定開發分支提交，再 fast-forward 到 main 觸發 Pages 部署。）
   回覆使用者可分享連結：
   `https://jackielin666.github.io/production-dashboard/reports/YYYY-MM.html`

## 期中預覽（月底廠務會議用）

1. 確認資料：`git pull` 取得使用者剛上傳的 `data/latest.xlsx`。
2. 產生預覽（不需 `--allow-incomplete`）：
   ```bash
   python3 tools/monthly_report.py --month YYYY-MM --preview
   ```
   預設以資料中該月最後一天為截止日；可加 `--asof YYYY-MM-DD` 指定。
3. 讀 `reports/YYYY-MM-preview.json`，寫 **1~2 段**精簡總評存成
   `reports/YYYY-MM-preview.narrative.txt`，重點放在「**開會時該問什麼**」：
   - 開頭註明「資料截至 M/D，與上月／去年同期比較，本月尚未結束」。
   - 同期比較的方向、`totals.projected` 推估全月量（說明僅供參考）。
   - A 級警示、`tracking` 中狀態為紅的事項、`dormant`（此時為「截至目前尚未生產」）中的主力品項
     → 建議會中向業務／生管確認是否已排程。
   - 資料 < 10 天時提醒波動大、結論保守。
   再執行：
   ```bash
   python3 tools/monthly_report.py --month YYYY-MM --preview --narrative reports/YYYY-MM-preview.narrative.txt
   ```
4. 提交、推送、回覆連結 `https://jackielin666.github.io/production-dashboard/reports/YYYY-MM-preview.html`。
   預覽每次產生都覆蓋同一檔；下月初產生正式月報後，索引頁自動改列正式版。

## 判定門檻（集中在 tools/report_config.json）

| 參數 | 預設 | 意義 |
|------|------|------|
| `decline_pct` | 10 | 降幅達此值即列入警示（比上月或去年同月）|
| `min_scale_pots` | 20 | 基準量門檻，排除零星小量品項的雜訊 |
| `top_n` | 10 | 圖表與清單取前幾名 |
| `dormant_lookback` / `dormant_min_active` | 6 / 3 | 前 6 個月中 ≥3 月有生產、本月掛零＝斷單 |
| `dormant_lost_months` | 3 | 連續未生產 ≥ 此月數 → 疑似流失 |
| `b_grade_pct` | 1 | B 級：非主力流失 ≥ 本業總鍋數的此百分比（約 30 鍋）|
| `red_core_loss_pct` | 15 | 紅燈：A 級（主力）流失合計 ≥ 本業總鍋數的此百分比 |
| `red_core_dormant_months` | 2 | 紅燈：主力品項連續斷單 ≥ 此月數 |
| `anomaly_*` | 見設定檔 | 資料異常（鍋數與重量不匹配）條件 |

使用者若要調整鬆緊，改這個檔即可，不要改寫程式邏輯。
完整規則說明見 `docs/monthly-report-logic.md`。

## 注意
- 報告數字必須可追溯：`tools/monthly_report.py` 的彙總規則與 `index.html` 的
  `buildDashboardData` 一致（同一份資料應得到同一組數字）。修改任一邊都要兩邊對齊。
- 同名但不同料號的品項會分別計算，報告中以料號區分，不要合併。
