---
name: monthly-report
description: 產生「生產鍋數月報」——依 data/latest.xlsx 分析指定月份的業務量（鍋數／重量）、接單下滑警示、斷單、成長品項與季節性，輸出 HTML 報告到 reports/ 並推上 GitHub Pages 取得可分享連結。當使用者說「幫我分析 2026年8月」「幫我出 8 月月報」「分析上個月」「月度分析」「月報」時使用。
---

# 生產鍋數月報

使用者每月初把上月完整的「產品製成率」Excel 覆蓋到 `data/latest.xlsx`，然後說
「幫我分析 YYYY年M月」，你就依本流程產出一份業務導向的月報並回覆連結。

**分析取向**：公司是業務導向、接單式生產。**生產鍋數直接反映業務量與營收** ——
某品項鍋數明顯下降 = 出貨量下降 = 接單量下滑 = 營收風險，主力品項尤其重要。
**製成率／品質不在本報告範圍**（由「產品工時及製成率統計系統」負責），不要加回來。

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
   - **整體業務量方向**：環比／同比／YTD，並用 `seasonal_note` 判斷是季節性還是實質衰退。
   - **營收風險**：`core_real`（主力品項實質衰退）與 `dormant`（斷單）是最該點名的，
     說明流失多少鍋數、可能的影響。
   - **注意 `verdict` 欄**：`batch` 代表鍋數降但重量持平＝批量調整，**不是接單下滑**，
     不要當成壞消息報上去。
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

## 判定門檻（集中在 tools/report_config.json）

| 參數 | 預設 | 意義 |
|------|------|------|
| `decline_pct` | 10 | 降幅達此值即列入警示（比上月或去年同月）|
| `min_scale_pots` | 20 | 基準量門檻，排除零星小量品項的雜訊 |
| `top_n` | 10 | 圖表與清單取前幾名 |
| `dormant_lookback` / `dormant_min_active` | 6 / 3 | 前 6 個月中 ≥3 月有生產、本月掛零＝斷單 |

使用者若要調整鬆緊，改這個檔即可，不要改寫程式邏輯。
完整規則說明見 `docs/monthly-report-logic.md`。

## 注意
- 報告數字必須可追溯：`tools/monthly_report.py` 的彙總規則與 `index.html` 的
  `buildDashboardData` 一致（同一份資料應得到同一組數字）。修改任一邊都要兩邊對齊。
- 同名但不同料號的品項會分別計算，報告中以料號區分，不要合併。
