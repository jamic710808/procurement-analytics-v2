# 開發紀錄 (Development Record)

## 專案標題：採購數據分析與戰略決策儀表板

### 開發目標
- 將原始單一平面文件（簡報、Excel數據）升級為 8 大維度的互動式資訊看板。
- 提供符合未來感與高層次決策 (Executive-level) 的 UI/UX 設計。
- 實踐被動報告轉為主動預警與趨勢預測（情境模擬器）。

### 採用技術棧
- **Frontend Core**: HTML5, CSS3, Vanilla JavaScript
- **Visual Design**: CSS Variables (Dark Theme), CSS Backdrop-filter (Glassmorphism)
- **Data Visualization**: Chart.js 3.x
- **Data Handling**: SheetJS (xlsx.js) 處理客端 Excel 解析
- **Icons**: FontAwesome 6 (CDN)
- **Data Generator**: Python 3, Pandas, Numpy

### 開發歷程與決策紀錄
1. **分析與洞察萃取階段**：
   - 梳理資料夾中的《現代化採購數據分析與戰略決策簡報》與 Excel 資料，提煉出 P2P 流程效率、供應商四象限、品質追蹤與成本控制等分析主題。
   - 撰寫 `analysis.md` 確立新世代設計建議與 8 大分頁架構。
2. **架構設計與基礎搭建**：
   - 設定深色調（`#0f172a`）與高對比色（藍、綠、紅）增加視覺重心。
   - 將所有前端檔案獨立為規範的架構 (`/css`, `/js`)，利於後續擴展模組。
3. **資料層與視覺層對接**：
   - 獨立出 `data_engine.js`，管理全局資料 (`AppData`)，模擬 Flux/Redux 的單一資訊來源概念。
   - `dashboard.js` 實現了多頁面無刷新切換 (SPA Router simulation)，並註冊多達 14 組以上的視覺圖表。
4. **資料連動測試**：
   - 撰寫 `generate_sample_data.py` 自動化產出多標籤頁 (Multisheet) Excel 測試資料。
   - 對接右上角的 File Uploader，將 FileReader 結果傳入 SheetJS，觸發介面 Refresh，驗證全域資料聯動。
5. **V2 儀表板架構升級與修復**：
   - **擴增頁面**：從 8 頁擴充至 10 頁，新增「合約生命週期追蹤」與「採購成熟度（戰略）地圖」。
   - **圖表補齊**：修復前期測試中空白圖表（例如：整體雷達圖、瀑布圖、各類表格、SPC 圖等）的渲染問題，確保畫面上每個數字皆有連動。
   - **匯入/匯出互動**：撰寫 `generate_sample.py` 動態產生 2024–2025 的 CSV 範例資料，並在上傳視窗中加入「下載範例資料檔」連結。全面改寫 JS 內部資料 Mapping 邏輯，確保使用者匯入 Excel/CSV 後，10 個頁面的圖表會立即重新計算（連動包含情境模擬器）。
   - **時效性更新**：介面預設濾波器及隨機資料產生器皆切換為 **2024-2025 年度**，契合最新業務分析需求。

### 結論
本次更新成功替換舊有靜態閱覽模式，賦予採購部門「自主查閱、重點提示、趨勢模擬」的數位化動能，並透過匯入/匯出機制的完善與時間軸推進，大幅降低資訊落差與溝通成本。
