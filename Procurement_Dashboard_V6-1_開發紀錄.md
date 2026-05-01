# 採購智能分析平台 V6.1 開發紀錄

**版本**：V6.1
**日期**：2026-03-22
**作者**：Claude Code AI

---

## 1. 版本比較總覽

| 項目 | V5 | V6 | V6.1 |
|------|:---:|:---:|:----:|
| 檔案大小 | 2650 行 | 2760 行 | 2765 行 |
| PDF 匯出方式 | 單頁 html2pdf | 多頁 html2canvas + jsPDF | 多頁 html2canvas + jsPDF |
| 企業級匯出 Modal | 無 | 有 | 有 |
| 主題偵測（匯出時） | 直接讀取 DOM | 依賴 checkbox 值 | 直接讀取 DOM |
| 主題切換按鈕 | 有 | 有 | 有 |

---

## 2. V5 → V6 重大改版差異

### 2.1 PDF 匯出架構重構

**V5 採用舊式 html2pdf.js**：
```javascript
// V5 程式碼（約第 2385 行）
function exportToPDF() {
  const currentTheme = document.body.getAttribute('data-theme');
  const isLight = currentTheme === 'light';
  // 使用 html2pdf.bundle 直接匯出
  html2pdf().set(opt).from(activePage).save();
}
```
- **限制**：僅能匯出當前可見頁面，無法批量匯出
- **優點**：實作簡單
- **缺點**：CSS 變數在半透明背景上常有渲染問題

**V6 採用新式 html2canvas + jsPDF**：
```javascript
// V6 程式碼（約第 2561 行）
async function executePDFExport(pagesToExport, options) {
  // 1. 使用 html2canvas 截圖每個頁面
  const canvas = html2canvas(pageNode, { scale: 2 });
  // 2. 將 canvas 轉為圖片後以 jsPDF 多頁組裝
  pdf.addImage(imgData, 'JPEG', margin, margin, usableW, usableH);
}
```
- **優點**：支援多頁批量匯出、浮水印、高品質模式、橫/直式切換
- **缺點**：實作複雜，需手動處理長頁面分割

### 2.2 企業級匯出 Modal

V6 新增 `showExportModal()` 函式，提供完整匯出設定介面：

| 功能 | 說明 |
|------|------|
| 頁面選擇 | 可勾選要匯出的頁面（全選/清除） |
| 浮水印 | 自訂文字浮水印 |
| 日期戳記 | 自動加入匯出時間 |
| 橫直版面 | 切換 A4 橫式/直式 |
| 高品質渲染 | 2x 像素密度（耗時較長） |

### 2.3 V6 原始碼結構

```
Procurement_Dashboard_V6.html
├── HTML 結構
│   ├── Sidebar 導航（11 頁面）
│   ├── Topbar 工具列
│   ├── Export Modal（第 562-602 行）
│   └── Page 內容區塊
│
├── JavaScript 函式
│   ├── nav() - 頁面導航
│   ├── renderPageCharts() - 圖表渲染
│   ├── showExportModal() - 顯示匯出設定 Modal
│   ├── proceedExportFromModal() - 解析匯出選項
│   └── executePDFExport() - 核心匯出引擎
│
└── CSS 樣式
    ├── :root 與 [data-theme="light"] 雙主題
    ├── 卡片、圖表、動畫等元件
    └── PDF 匯出專用 solid 樣式
```

---

## 3. V6 → V6.1 修復紀錄

### 3.1 問題描述

**V6 PDF 匯出失敗**：在某些情況下，PDF 匯出後的主題顏色與預期不符（亮色主題匯出成深色，或反之）。

### 3.2 根本原因分析

V6 在計算主題時使用了错误的偵測方式：

```javascript
// V6 原始程式碼（第 2554 行）
const useLight = !options.darkMode;
```

`options.darkMode` 來自於 Modal 中的 checkbox（`exp-dark-mode`），但該 checkbox 被設為 `disabled`，其值為 `!isLightTheme`（亮色主題時為 false）。問題在於：

1. 用戶可能在匯出期間點擊主題切換按鈕
2. checkbox 狀態與 DOM 實際主題可能不同步
3. 導致 `useLight` 判斷錯誤，CSS 注入錯誤的主題樣式

### 3.3 解決方案

**V6-1 修復方式**：直接讀取 DOM 實際主題狀態

```javascript
// V6-1 修復後（第 2554 行）
const useLight = document.body.getAttribute('data-theme') === 'light';
```

此方式直接讀取 `<body data-theme="light">` 屬性，確保與視覺呈現完全一致。

---

## 4. V6.1 技術亮點

### 4.1 雙 CDN jsPDF 支援

為了解決 jsPDF CDN 載入順序問題，V6 實作雙重 fallback：

```javascript
const jsPDFCtor = (window.jspdf && window.jspdf.jsPDF)
  ? window.jspdf.jsPDF
  : (window.jsPDF || null);
```

### 4.2 長頁面智慧分割

當頁面高度超過 A4 可用高度時，系統會自動分割：

```javascript
const ratio = usableW / srcW;
const sliceHeightPx = Math.max(1, Math.floor(usableH / ratio));
let sy = 0;
while (sy < srcH) {
  const sh = Math.min(sliceHeightPx, srcH - sy);
  // ... 裁切並新增至 PDF
  sy += sh;
}
```

### 4.3 匯出期間主題鎖定

為確保匯出期間主題穩定，V6 實作了完整的状态保存/恢復機制：

```javascript
// 保存原始狀態
const originalTheme = document.body.getAttribute('data-theme');
const activePageBeforeExport = activeNav ? activeNav.dataset.page : null;

// 匯出完成後恢復
document.body.setAttribute('data-theme', originalTheme);
nav(activePageBeforeExport);
```

### 4.4 跨主題 PDF 樣式注入

針對亮色/深色主題，系統在匯出時注入對應的實色 CSS：

```javascript
const exportStyle = document.createElement('style');
exportStyle.textContent = useLight ? `
  #pdf-export-solid { background: #f9f6f0 !important; color: #33312e !important; }
  #pdf-export-solid .card { background: #ffffff !important; }
  /* ... 亮色主題完整覆寫 */
` : `
  #pdf-export-solid { background: #0f172a !important; color: #f8fafc !important; }
  /* ... 深色主題完整覆寫 */
`;
```

---

## 5. 已知限制

| 限制 | 說明 | 建議 |
|------|------|------|
| backdrop-filter 無法渲染 | html2canvas 不支援 CSS `backdrop-filter` | 匯出前系統自動清除此效果 |
| 外部字體依賴網路 | Font Awesome 等 CDN 字體需連線 | 確保匯出時網路連線正常 |
| 動畫元素定格 | Chart.js 動畫在匯出時停用 | 確保圖表已完全渲染 |
| 浮水印僅文字 | 目前僅支援純文字浮水印 | 未來可擴展支援圖片浮水印 |

---

## 6. 測試要點

### 6.1 主題切換測試
1. 開啟亮色主題 → 匯出 PDF → 確認 PDF 為亮色
2. 開啟深色主題 → 匯出 PDF → 確認 PDF 為深色
3. 在深色主題下，點擊主題切換（快速）→ 立即匯出 → 確認主題正確

### 6.2 多頁選擇測試
1. 只選 1 頁 → 匯出 → 確認僅 1 頁
2. 全選 → 匯出 → 確認所有頁面
3. 取消全選 → 匯出 → 確認提示「至少需選擇一頁」

### 6.3 品質測試
1. 高品質模式（2x）→ 匯出 → 確認解析度
2. 標準品質 → 匯出 → 確認效能提升

---

## 7. 未來改進方向

- [ ] 支援圖片浮水印
- [ ] 支援頁面範本（公司 Logo、頁尾資訊）
- [ ] 預覽功能（匯出前先預覽效果）
- [ ] Excel 匯出（除了 PDF 外）
- [ ] 自動化排程匯出

---

**文件版本**：1.0
**最後更新**：2026-03-22
**狀態**：已完成
