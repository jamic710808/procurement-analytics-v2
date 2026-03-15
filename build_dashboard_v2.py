import os
import io

OUTPUT_FILE = r'c:\Users\jamic\採購分析\採購分析V2_Dashboard.html'

HTML_HEAD_CSS = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>採購分析 V2 智能看板</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#0f172a;--surface:#1e293b;--glass:rgba(30,41,59,0.7);--border:rgba(255,255,255,0.08);
  --primary:#6366f1;--accent:#0ea5e9;--gold:#f59e0b;--red:#ef4444;--green:#10b981;
  --text:#f8fafc;--muted:#94a3b8;--radius:12px;--rad-lg:16px;
}
*{margin:0;padding:0;box-sizing:border-box;font-family:'Segoe UI',system-ui,sans-serif;}
body{background:var(--bg);color:var(--text);display:flex;height:100vh;overflow:hidden;}

/* Navbar / Sidebar */
.sidebar{width:220px;background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;flex-shrink:0;}
.logo{padding:20px 16px;font-size:15px;font-weight:700;display:flex;align-items:center;gap:10px;border-bottom:1px solid var(--border);}
.logo i{color:var(--accent);font-size:22px;}
.nav-group{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;padding:16px 16px 8px;}
.nav-item{padding:10px 16px;cursor:pointer;display:flex;align-items:center;gap:10px;font-size:13px;color:var(--muted);border-left:3px solid transparent;transition:0.2s;}
.nav-item:hover{background:rgba(255,255,255,0.03);color:var(--text);}
.nav-item.active{background:rgba(99,102,241,0.1);color:var(--primary);border-left-color:var(--primary);font-weight:600;}
.nav-item i{width:16px;text-align:center;}

/* Main Area */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;}
.topbar{height:60px;background:var(--glass);backdrop-filter:blur(12px);border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;padding:0 24px;}
.filters{display:flex;gap:12px;align-items:center;}
.filters select{background:#0f172a;color:var(--text);border:1px solid var(--border);padding:6px 12px;border-radius:6px;font-size:12px;cursor:pointer;outline:none;}
.actions .btn{padding:6px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:none;display:inline-flex;align-items:center;gap:6px;}
.btn-primary{background:var(--primary);color:#fff;}
.btn-outline{background:transparent;border:1px solid var(--border);color:var(--text);}
.btn:hover{opacity:0.9;transform:translateY(-1px);}

/* Content */
.content{flex:1;overflow-y:auto;padding:24px;}
.page{display:none;animation:fadeIn 0.3s ease;}
.page.active{display:block;}
@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
.pg-title{font-size:20px;font-weight:700;margin-bottom:6px;display:flex;align-items:center;gap:8px;}
.pg-sub{font-size:13px;color:var(--muted);margin-bottom:20px;}

/* Grid System */
.grid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:16px;}
.grid-3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-bottom:16px;}
.grid-2{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;margin-bottom:16px;}
.grid-12{display:grid;grid-template-columns:1fr 2fr;gap:16px;margin-bottom:16px;}
.grid-21{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:16px;}

/* Cards & KPIs */
.card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px;box-shadow:0 4px 6px rgba(0,0,0,0.2);}
.card-header{font-size:13px;font-weight:600;color:var(--muted);margin-bottom:12px;display:flex;justify-content:space-between;align-items:center;}
.kpi-val{font-size:28px;font-weight:800;color:var(--text);margin:4px 0;}
.kpi-delta{font-size:12px;display:inline-flex;align-items:center;gap:4px;}
.up{color:var(--green);} .down{color:var(--red);} .warn{color:var(--gold);}
.chart-container{position:relative;height:240px;width:100%;}
.chart-lg{height:320px;}

/* Tables */
.table-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{text-align:left;padding:10px 12px;background:rgba(255,255,255,0.03);color:var(--muted);font-weight:600;border-bottom:1px solid var(--border);}
td{padding:10px 12px;border-bottom:1px solid var(--border);color:var(--text);}
tr:hover td{background:rgba(255,255,255,0.02);}
.badge{padding:3px 8px;border-radius:999px;font-size:10px;font-weight:600;}
.b-green{background:rgba(16,185,129,0.2);color:var(--green);}
.b-red{background:rgba(239,68,68,0.2);color:var(--red);}
.b-gold{background:rgba(245,158,11,0.2);color:var(--gold);}
.b-blue{background:rgba(14,165,233,0.2);color:var(--accent);}

/* Utilities */
.flex-between{display:flex;justify-content:space-between;align-items:center;}
.alert-box{background:rgba(239,68,68,0.1);border-left:3px solid var(--red);padding:12px;border-radius:6px;margin-bottom:12px;display:flex;gap:12px;align-items:flex-start;font-size:12px;}
.alert-title{font-weight:600;color:var(--text);margin-bottom:4px;}

/* Slider */
.slider-group{margin-bottom:16px;}
.slider-group label{display:flex;justify-content:space-between;font-size:12px;margin-bottom:6px;color:var(--text);}
input[type=range]{width:100%;accent-color:var(--primary);}

/* Modal Overlay */
.overlay{position:fixed;inset:0;background:rgba(0,0,0,0.8);backdrop-filter:blur(4px);z-index:1000;display:none;place-items:center;}
.overlay.show{display:grid;}
.modal{background:var(--surface);width:500px;border-radius:var(--rad-lg);border:1px solid var(--border);padding:24px;box-shadow:0 10px 25px rgba(0,0,0,0.5);}
.modal-title{font-size:18px;font-weight:700;margin-bottom:16px;color:var(--accent);}
.file-drop{border:2px dashed var(--muted);padding:40px;text-align:center;border-radius:var(--radius);color:var(--muted);cursor:pointer;transition:0.2s;margin-bottom:16px;}
.file-drop:hover{border-color:var(--primary);background:rgba(99,102,241,0.05);color:var(--text);}
</style>
</head>
<body>
"""

HTML_BODY = """
<div class="sidebar">
  <div class="logo"><i class="fa-solid fa-layer-group"></i> 採購智能平台 V2</div>
  <div style="flex:1;overflow-y:auto;padding-top:10px;">
    <div class="nav-group">核心分析 (Core)</div>
    <div class="nav-item active" onclick="nav('p-summary')"><i class="fa-solid fa-gauge"></i>決策總覽</div>
    <div class="nav-item" onclick="nav('p-spend')"><i class="fa-solid fa-chart-pie"></i>支出分析</div>
    <div class="nav-item" onclick="nav('p-supplier')"><i class="fa-solid fa-users"></i>供應商矩陣</div>
    <div class="nav-item" onclick="nav('p-cost')"><i class="fa-solid fa-piggy-bank"></i>成本節約</div>
    <div class="nav-item" onclick="nav('p-risk')"><i class="fa-solid fa-shield-halved"></i>風險管控</div>
    
    <div class="nav-group">進階模組 (Advanced)</div>
    <div class="nav-item" onclick="nav('p-esg')"><i class="fa-solid fa-leaf"></i>ESG 永續評核</div>
    <div class="nav-item" onclick="nav('p-spc')"><i class="fa-solid fa-chart-line"></i>品質 SPC 管制</div>
    <div class="nav-item" onclick="nav('p-contract')"><i class="fa-solid fa-file-signature"></i>合約生命週期</div>
    <div class="nav-item" onclick="nav('p-maturity')"><i class="fa-solid fa-graduation-cap"></i>採購成熟度</div>
    <div class="nav-item" onclick="nav('p-sim')"><i class="fa-solid fa-sliders"></i>情境模擬站</div>
  </div>
  <div style="padding:16px;border-top:1px solid var(--border);font-size:10px;color:var(--muted);text-align:center;">
    v2.0.0 | 即時數據連線正常
  </div>
</div>

<div class="main">
  <div class="topbar">
    <div class="filters">
      <span style="font-size:12px;color:var(--muted)"><i class="fa-solid fa-filter"></i> 篩選:</span>
      <select id="flt-year" onchange="applyFilters()">
        <option value="all">所有年份 (All Years)</option>
        <option value="2023">2023</option>
        <option value="2024">2024</option>
      </select>
      <select id="flt-cat" onchange="applyFilters()">
        <option value="all">所有品類 (All Categories)</option>
        <option value="電子零組件">電子零組件</option>
        <option value="機構件">機構件</option>
        <option value="包材">包材</option>
        <option value="間接物料">間接物料</option>
      </select>
    </div>
    <div class="actions">
      <button class="btn btn-outline" style="margin-right:8px;" onclick="generateDummyData()"><i class="fa-solid fa-arrows-rotate"></i> 隨機再生資料</button>
      <button class="btn btn-primary" onclick="showUpload()"><i class="fa-solid fa-cloud-arrow-up"></i> 匯入 XLSX</button>
    </div>
  </div>

  <div class="content">
    
    <!-- P1: SUMMARY -->
    <div id="p-summary" class="page active">
      <div class="pg-title">決策總覽 (Executive Dashboard)</div>
      <div class="pg-sub">整體採購概況與關鍵績效指標即時監控</div>
      
      <div class="grid-4">
        <div class="card">
          <div class="card-header">總採購支出</div>
          <div class="kpi-val" id="k-spend">$0</div>
          <div class="kpi-delta up"><i class="fa-solid fa-arrow-up"></i> <span>即時計算</span></div>
        </div>
        <div class="card">
          <div class="card-header">成本節約額</div>
          <div class="kpi-val" id="k-save" style="color:var(--green)">$0</div>
          <div class="kpi-delta up"><i class="fa-solid fa-check"></i> <span>達標率 94%</span></div>
        </div>
        <div class="card">
          <div class="card-header">活躍供應商</div>
          <div class="kpi-val" id="k-sup">0</div>
          <div class="kpi-delta warn"><i class="fa-solid fa-minus"></i> <span>集中度過高警示</span></div>
        </div>
        <div class="card">
          <div class="card-header">平均交期達成率</div>
          <div class="kpi-val" id="k-otd" style="color:var(--accent)">0%</div>
          <div class="kpi-delta up"><i class="fa-solid fa-arrow-up"></i> <span>較上月 +1.2%</span></div>
        </div>
      </div>
      
      <div class="grid-12">
        <div class="card">
          <div class="card-header">即時警示通知</div>
          <div id="alert-container" style="max-height:240px;overflow-y:auto;">
            <!-- auto generated -->
          </div>
        </div>
        <div class="card">
          <div class="card-header">年度月度支出趨勢 (YTD Spend Trend)</div>
          <div class="chart-container"><canvas id="c-trend"></canvas></div>
        </div>
      </div>
      
      <div class="grid-3">
        <div class="card">
          <div class="card-header">品類支出佔比</div>
          <div class="chart-container"><canvas id="c-cat-donut"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">供應商風險分佈</div>
          <div class="chart-container"><canvas id="c-risk-donut"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">總結目標達成狀況</div>
          <div class="chart-container"><canvas id="c-radar-ov"></canvas></div>
        </div>
      </div>
    </div>

    <!-- P2: SPEND -->
    <div id="p-spend" class="page">
      <div class="pg-title">支出分析 (Spend Analytics)</div>
      <div class="pg-sub">品類、部門與供應商支出結構下鑽分析</div>
      
      <div class="grid-2">
        <div class="card">
          <div class="card-header">各部門採購金額排行</div>
          <div class="chart-container chart-lg"><canvas id="c-spend-dept"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">各品類支出分佈 (Waterfall)</div>
          <div class="chart-container chart-lg"><canvas id="c-spend-waterfall"></canvas></div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">Top 10 供應商支出總覽</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>供應商名稱</th><th>提供品類</th><th>支出金額 (TWD)</th><th>訂單數</th><th>平均單價區間</th></tr></thead>
            <tbody id="tb-spend-sup"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- P3: SUPPLIER MATRIX -->
    <div id="p-supplier" class="page">
      <div class="pg-title">供應商矩陣 (Supplier Matrix)</div>
      <div class="pg-sub">Kraljic 矩陣與關鍵績效評量</div>
      
      <div class="grid-12">
        <div class="card">
          <div class="card-header">績效金字塔</div>
          <div class="chart-container chart-lg"><canvas id="c-sup-radar"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">供應商象限分析 (Bubble Matrix - 品質 vs 交期)</div>
          <div class="chart-container chart-lg"><canvas id="c-sup-bubble"></canvas></div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">供應商績效排行榜</div>
        <div class="table-wrap">
          <table>
            <thead><tr><th>排名</th><th>供應商</th><th>綜合得分</th><th>品質良率</th><th>交期準確率</th><th>ESG評級</th><th>狀態建議</th></tr></thead>
            <tbody id="tb-sup-rank"></tbody>
          </table>
        </div>
      </div>
    </div>
    
    <!-- P4: COST -->
    <div id="p-cost" class="page">
      <div class="pg-title">成本節約 (Cost Saving)</div>
      <div class="grid-21">
        <div class="card">
          <div class="card-header">月度節約金額趨勢</div>
          <div class="chart-container chart-lg"><canvas id="c-cost-trend"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">節約來源佔比</div>
          <div class="chart-container chart-lg"><canvas id="c-cost-pie"></canvas></div>
        </div>
      </div>
    </div>

    <!-- P5: RISK -->
    <div id="p-risk" class="page">
      <div class="pg-title">風險管控 (Risk Management)</div>
      <div class="grid-3">
        <div class="card">
          <div class="card-header">地緣風險集中度</div>
          <div class="chart-container"><canvas id="c-risk-geo"></canvas></div>
        </div>
        <div class="card" style="grid-column: span 2">
          <div class="card-header">高風險事件追蹤表</div>
          <div class="table-wrap" style="height:240px;overflow-y:auto">
            <table>
              <thead><tr><th>供應商</th><th>風險類型</th><th>風險指數</th><th>衝擊預估</th><th>改善措施</th></tr></thead>
              <tbody id="tb-risk"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
    
    <!-- P6: ESG -->
    <div id="p-esg" class="page">
      <div class="pg-title">ESG 永續評核 (ESG Analytics)</div>
      <div class="grid-2">
        <div class="card">
          <div class="card-header">整體 E/S/G 評分分佈</div>
          <div class="chart-container chart-lg"><canvas id="c-esg-bar"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">碳排放量追蹤 (Scope 3)</div>
          <div class="chart-container chart-lg"><canvas id="c-esg-carbon"></canvas></div>
        </div>
      </div>
    </div>
    
    <!-- P7: SPC -->
    <div id="p-spc" class="page">
      <div class="pg-title">品質 SPC 管制 (Statistical Process Control)</div>
      <div class="pg-sub">進料品質檢驗 (IQC) X-Bar 與 R 管制圖</div>
      <div class="card" style="margin-bottom:16px;">
        <div class="card-header">X-Bar Chart (平均值管制圖)</div>
        <div class="chart-container"><canvas id="c-spc-x"></canvas></div>
      </div>
      <div class="card">
        <div class="card-header">異常批次處置清單</div>
        <div class="table-wrap">
          <table id="tb-spc">
            <thead><tr><th>批次號</th><th>檢驗日期</th><th>量測均值</th><th>UCL</th><th>LCL</th><th>判定</th></tr></thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
    
    <!-- P8: CONTRACT -->
    <div id="p-contract" class="page">
      <div class="pg-title">合約生命週期 (Contract Management)</div>
      <div class="card">
        <div class="card-header">合約到期預警雷達 (依品類)</div>
        <div class="chart-container chart-lg"><canvas id="c-contract-bar"></canvas></div>
      </div>
    </div>
    
    <!-- P9: MATURITY -->
    <div id="p-maturity" class="page">
      <div class="pg-title">採購成熟度 (Procurement Maturity)</div>
      <div class="grid-2">
        <div class="card">
          <div class="card-header">五維度成熟度模型</div>
          <div class="chart-container chart-lg"><canvas id="c-mat-radar"></canvas></div>
        </div>
        <div class="card">
          <div class="card-header">改善計畫路徑圖</div>
          <div style="padding:20px;color:var(--muted);line-height:1.6;font-size:13px;">
            <p><i class="fa-solid fa-check" style="color:var(--green)"></i> 第一階段：無紙化與自動對帳 (100% 達成)</p>
            <p><i class="fa-solid fa-spinner" style="color:var(--accent)"></i> 第二階段：供應商數位協同平台 (進行中 65%)</p>
            <p><i class="fa-solid fa-lock" style="color:var(--muted)"></i> 第三階段：AI 預測性採購模型 (計畫中 Q4)</p>
            <p><i class="fa-solid fa-lock" style="color:var(--muted)"></i> 第四階段：Scope 3 碳盤查自動串接 (2025 計畫)</p>
          </div>
        </div>
      </div>
    </div>

    <!-- P10: SIM -->
    <div id="p-sim" class="page">
      <div class="pg-title">情境模擬站 (Scenario Simulator)</div>
      <div class="grid-12">
        <div class="card">
          <div class="card-header">參數調整 (Monte Carlo 變數)</div>
          <div class="slider-group">
            <label><span>原物料價格漲幅</span><span id="v-raw">0%</span></label>
            <input type="range" id="sl-raw" min="-20" max="50" value="0" oninput="runSim()">
          </div>
          <div class="slider-group">
            <label><span>匯率波動 (USD/TWD)</span><span id="v-fx">32.0</span></label>
            <input type="range" id="sl-fx" min="28" max="36" step="0.1" value="32.0" oninput="runSim()">
          </div>
          <div class="slider-group">
            <label><span>關稅與運費變化</span><span id="v-freight">0%</span></label>
            <input type="range" id="sl-freight" min="-10" max="100" value="0" oninput="runSim()">
          </div>
          <button class="btn btn-primary" style="width:100%;justify-content:center;margin-top:10px;" onclick="runSim()">重新計算推演</button>
        </div>
        <div class="card">
          <div class="card-header">預估採購總成本影響 (TCO Impact)</div>
          <div class="chart-container chart-lg"><canvas id="c-sim"></canvas></div>
        </div>
      </div>
    </div>

  </div>
</div>

<!-- Upload Modal -->
<div class="overlay" id="uploadModal">
  <div class="modal">
    <div class="flex-between" style="margin-bottom:16px;">
      <div class="modal-title">匯入更新資料</div>
      <button class="btn btn-outline" style="border:none;font-size:18px;" onclick="closeUpload()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div class="file-drop" id="drop-zone">
      <i class="fa-solid fa-cloud-arrow-up" style="font-size:32px;margin-bottom:12px;"></i>
      <p>點擊上傳或拖曳 Excel (XLSX) 檔案至此</p>
      <input type="file" id="fileInput" style="display:none" accept=".xlsx, .xls, .csv">
    </div>
    <p style="font-size:12px;color:var(--muted);margin-bottom:8px;">支援欄位要求：</p>
    <div style="font-size:11px;color:var(--accent);background:rgba(14,165,233,0.1);padding:10px;border-radius:6px;line-height:1.5;">
      [日期] [供應商] [品類] [部門] [採購金額] [品質分數] [交期分數] [ESG評分] [風險係數]
    </div>
  </div>
</div>
"""

HTML_JS = """
<script>
// ==== Core Data Engine ====
let rawData = [];
let fData = [];

// Globals for charts
let charts = {};

function initApp() {
    generateDummyData();
    initCharts();
    setupEventListeners();
    applyFilters();
}

function nav(pageId) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(p => p.classList.remove('active'));
    document.getElementById(pageId).classList.add('active');
    event.currentTarget.classList.add('active');
    
    // Trigger resize to fix chart.js rendering bug on hidden tabs
    window.dispatchEvent(new Event('resize'));
}

// Data generator
function generateDummyData() {
    console.log("Generating dummy data...");
    rawData = [];
    const suppliers = ['TSMC (晶圓)', 'Foxconn (組裝)', 'Delta (電源)', 'Yageo (被動元件)', 'Cisco (網通)', 'AWS (雲服務)', 'Logitech (周邊)'];
    const categories = ['電子零組件', '機構件', '包材', '間接物料'];
    const depts = ['製造部', '研發部', '營運部', '行銷部'];
    
    for(let i=1; i<=300; i++) {
        // Random date between 2023-01-01 and 2024-12-31
        let start = new Date(2023, 0, 1).getTime();
        let end = new Date(2024, 11, 31).getTime();
        let date = new Date(start + Math.random() * (end - start));
        
        let spend = Math.floor(Math.random() * 8000000) + 100000;
        let p_save = Math.random() * 0.15; // 0~15% saving
        let quality = Math.floor(Math.random() * 20) + 80; // 80~100
        let otd = Math.floor(Math.random() * 25) + 75; // 75~100
        let risk = Math.floor(Math.random() * 60) + 10; // 10~70
        let esg = Math.floor(Math.random() * 40) + 50; // 50~90
        
        rawData.push({
            id: 'PO-202' + i.toString().padStart(3,'0'),
            date: date.toISOString().split('T')[0],
            year: date.getFullYear().toString(),
            month: date.getMonth() + 1,
            supplier: suppliers[i % suppliers.length],
            category: categories[Math.floor(Math.random()*categories.length)],
            dept: depts[Math.floor(Math.random()*depts.length)],
            spend: spend,
            saved: spend * p_save,
            quality: quality,
            otd: otd,
            risk: risk,
            esg: esg
        });
    }
    
    // Check if UI is ready
    if(document.getElementById('k-spend')) {
      applyFilters();
    }
}

function applyFilters() {
    let yr = document.getElementById('flt-year').value;
    let cat = document.getElementById('flt-cat').value;
    
    fData = rawData.filter(d => {
        let yMatch = (yr === 'all') || (d.year === yr);
        let cMatch = (cat === 'all') || (d.category === cat);
        return yMatch && cMatch;
    });
    
    updateKPIs();
    updateCharts();
}

function updateKPIs() {
    let totalSpend = fData.reduce((acc, curr) => acc + curr.spend, 0);
    let totalSaved = fData.reduce((acc, curr) => acc + curr.saved, 0);
    let uniqSup = new Set(fData.map(d=>d.supplier)).size;
    let avgOtd = fData.reduce((acc, curr) => acc + curr.otd, 0) / (fData.length || 1);
    
    document.getElementById('k-spend').innerText = '$' + (totalSpend/1000000).toFixed(2) + 'M';
    document.getElementById('k-save').innerText = '$' + (totalSaved/1000000).toFixed(2) + 'M';
    document.getElementById('k-sup').innerText = uniqSup;
    document.getElementById('k-otd').innerText = avgOtd.toFixed(1) + '%';
    
    generateAlerts();
}

function generateAlerts() {
    let container = document.getElementById('alert-container');
    container.innerHTML = '';
    
    // Find terrible OTD
    let badOTD = fData.filter(d => d.otd < 80).slice(0,3);
    badOTD.forEach(x => {
        container.innerHTML += `<div class="alert-box">
            <i class="fa-solid fa-triangle-exclamation" style="font-size:16px;color:var(--red);margin-top:2px;"></i>
            <div>
              <div class="alert-title">交期嚴重延誤警告</div>
              <div style="color:var(--muted)">供應商 <b>${x.supplier}</b> 在單號 ${x.id} 的交期僅達成 ${x.otd}%，嚴重影響 ${x.dept} 排程。</div>
            </div>
        </div>`;
    });
}

function initCharts() {
    Chart.defaults.color = '#94a3b8';
    Chart.defaults.font.family = "'Segoe UI', sans-serif";
    Chart.defaults.borderColor = 'rgba(255,255,255,0.05)';
    
    const tooltipOptions = {
        backgroundColor: 'rgba(15, 23, 42, 0.9)',
        titleColor: '#fff',
        bodyColor: '#e2e8f0',
        borderColor: 'rgba(255,255,255,0.1)',
        borderWidth: 1,
        padding: 10
    };

    // Theme colors
    const colors = ['#6366f1', '#0ea5e9', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899'];

    // 1. Trend Line Chart
    charts.trend = new Chart(document.getElementById('c-trend'), {
        type: 'line',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: tooltipOptions }, elements: { line: { tension: 0.4 } } }
    });

    // 2. Category Donut
    charts.cat = new Chart(document.getElementById('c-cat-donut'), {
        type: 'doughnut',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: { legend: { position: 'right' }, tooltip: tooltipOptions } }
    });
    
    // 3. Risk Donut
    charts.risk = new Chart(document.getElementById('c-risk-donut'), {
        type: 'doughnut',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, cutout: '70%', plugins: { legend: { position: 'right' }, tooltip: tooltipOptions } }
    });
    
    // 4. Radar (Summary)
    charts.radarOv = new Chart(document.getElementById('c-radar-ov'), {
        type: 'radar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, scales: { r: { grid: { color: 'rgba(255,255,255,0.1)' }, angleLines: { color: 'rgba(255,255,255,0.1)' }, ticks: { display: false } } }, plugins: { legend: { display: false } } }
    });

    // 5. Spend by Dept (Bar)
    charts.spendDept = new Chart(document.getElementById('c-spend-dept'), {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display:false } } }
    });
    
    // 6. Spend Waterfall (Mocked using Bar chart with customized logic for visual)
    charts.waterfall = new Chart(document.getElementById('c-spend-waterfall'), {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 7. Supplier Bubble Matrix
    charts.supBubble = new Chart(document.getElementById('c-sup-bubble'), {
        type: 'bubble',
        data: { datasets: [] },
        options: { 
            responsive: true, maintainAspectRatio: false,
            scales: { x: { title: { display: true, text: '品質良率 (%)' } }, y: { title: { display: true, text: '交期準確率 (%)' } } },
            plugins: { tooltip: tooltipOptions } 
        }
    });

    // 8. Cost Trend
    charts.costTrend = new Chart(document.getElementById('c-cost-trend'), {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 9. Cost Pie
    charts.costPie = new Chart(document.getElementById('c-cost-pie'), {
        type: 'pie',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 10. Geo Risk
    charts.riskGeo = new Chart(document.getElementById('c-risk-geo'), {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y' }
    });

    // 11. ESG Bar
    charts.esgBar = new Chart(document.getElementById('c-esg-bar'), {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 12. ESG Carbon Line
    charts.esgCarbon = new Chart(document.getElementById('c-esg-carbon'), {
        type: 'line',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // 13. SPC X-Bar
    charts.spcX = new Chart(document.getElementById('c-spc-x'), {
        type: 'line',
        data: { labels: [], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false, plugins: { annotation: {} } }
    });

    // 14. Simulator
    charts.sim = new Chart(document.getElementById('c-sim'), {
        type: 'bar',
        data: { labels: ['基準情境 (Base)', '推演情境 (Simulated)'], datasets: [] },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

function updateCharts() {
    if(!charts.trend) return;
    const colors = ['#6366f1', '#0ea5e9', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899'];

    // Group by month
    let months = [...new Set(fData.map(d=>d.month))].sort((a,b)=>a-b);
    let spendByMonth = months.map(m => fData.filter(d=>d.month===m).reduce((a,c)=>a+c.spend,0));
    charts.trend.data = {
        labels: months.map(m => `Month ${m}`),
        datasets: [{
            label: 'Total Spend',
            data: spendByMonth,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99,102,241,0.1)',
            fill: true
        }]
    };
    charts.trend.update();

    // Group by Category
    let cats = [...new Set(fData.map(d=>d.category))];
    let spendByCat = cats.map(c => fData.filter(d=>d.category===c).reduce((a,x)=>a+x.spend,0));
    charts.cat.data = {
        labels: cats,
        datasets: [{ data: spendByCat, backgroundColor: colors }]
    };
    charts.cat.update();

    // Risk Doughnut
    let riskLevels = [
        {l: '高風險(>40)', v: fData.filter(d=>d.risk>40).length},
        {l: '中風險(20-40)', v: fData.filter(d=>d.risk>20 && d.risk<=40).length},
        {l: '低風險(<20)', v: fData.filter(d=>d.risk<=20).length}
    ];
    charts.risk.data = {
        labels: riskLevels.map(r=>r.l),
        datasets: [{ data: riskLevels.map(r=>r.v), backgroundColor: ['#ef4444', '#f59e0b', '#10b981'] }]
    };
    charts.risk.update();

    // Spend by Dept
    let depts = [...new Set(fData.map(d=>d.dept))];
    let spendByDept = depts.map(dp => fData.filter(d=>d.dept===dp).reduce((a,x)=>a+x.spend,0));
    charts.spendDept.data = {
        labels: depts,
        datasets: [{ label: '金額', data: spendByDept, backgroundColor: '#0ea5e9' }]
    };
    charts.spendDept.update();

    // Bubble Matrix
    let supAggr = {};
    fData.forEach(d => {
        if(!supAggr[d.supplier]) supAggr[d.supplier] = {s:0, q:0, o:0, c:0};
        supAggr[d.supplier].s += d.spend;
        supAggr[d.supplier].q += d.quality;
        supAggr[d.supplier].o += d.otd;
        supAggr[d.supplier].c += 1;
    });
    
    let bubbleData = Object.keys(supAggr).map((sup, idx) => ({
        label: sup,
        data: [{
            x: supAggr[sup].q / supAggr[sup].c,
            y: supAggr[sup].o / supAggr[sup].c,
            r: Math.max(8, (supAggr[sup].s / 10000000)) // Size by spend
        }],
        backgroundColor: colors[idx % colors.length]
    }));
    charts.supBubble.data = { datasets: bubbleData };
    charts.supBubble.update();

    // Populate Top Suppliers Table
    let sortedSups = Object.keys(supAggr).map(sup => ({
        name: sup,
        spend: supAggr[sup].s,
        orders: supAggr[sup].c,
        avgQ: (supAggr[sup].q / supAggr[sup].c).toFixed(1)
    })).sort((a,b)=>b.spend - a.spend).slice(0, 10);
    
    let tbHtml = '';
    sortedSups.forEach(s => {
        tbHtml += `<tr>
            <td><b>${s.name}</b></td>
            <td><span class="badge b-blue">主要</span></td>
            <td>$${(s.spend/1000000).toFixed(2)}M</td>
            <td>${s.orders}</td>
            <td>${s.avgQ}% (品質)</td>
        </tr>`;
    });
    let tbel = document.getElementById('tb-spend-sup');
    if(tbel) tbel.innerHTML = tbHtml;

    // Run Simulator visually once
    if(typeof runSim === 'function') runSim();
}

function runSim() {
    let rawDelta = parseFloat(document.getElementById('sl-raw').value);
    let fxVal = parseFloat(document.getElementById('sl-fx').value);
    let freightDelta = parseFloat(document.getElementById('sl-freight').value);
    
    document.getElementById('v-raw').innerText = rawDelta + '%';
    document.getElementById('v-fx').innerText = fxVal;
    document.getElementById('v-freight').innerText = freightDelta + '%';
    
    let baseTotal = fData.reduce((a,c)=>a+c.spend,0);
    
    // Naive formula: base * (1 + rawDelta/100 * 0.6 + fx factor * 0.3 + freight/100 * 0.1)
    let fxFactor = (fxVal - 32.0) / 32.0; 
    let multiplier = 1 + (rawDelta/100)*0.6 + (fxFactor)*0.3 + (freightDelta/100)*0.1;
    let simTotal = baseTotal * multiplier;
    
    if(charts.sim) {
        charts.sim.data = {
            labels: ['基準情境 (Base)', '推演情境 (Simulated)'],
            datasets: [{
                label: '預估總成本',
                data: [baseTotal, simTotal],
                backgroundColor: ['#6366f1', (simTotal > baseTotal ? '#ef4444' : '#10b981')]
            }]
        };
        charts.sim.update();
    }
}

// Upload Data Logic
function showUpload() { document.getElementById('uploadModal').classList.add('show'); }
function closeUpload() { document.getElementById('uploadModal').classList.remove('show'); }

function setupEventListeners() {
    let dropZone = document.getElementById('drop-zone');
    let fileInput = document.getElementById('fileInput');
    
    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.style.borderColor = '#0ea5e9'; });
    dropZone.addEventListener('dragleave', (e) => { e.preventDefault(); dropZone.style.borderColor = '#94a3b8'; });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.style.borderColor = '#94a3b8';
        if(e.dataTransfer.files.length > 0) processFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', (e) => {
        if(e.target.files.length > 0) processFile(e.target.files[0]);
    });
}

function processFile(file) {
    // Uses SheetJS to read file
    let reader = new FileReader();
    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, {type: 'array'});
        let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        let excelData = XLSX.utils.sheet_to_json(firstSheet);
        
        if(excelData.length > 0) {
            alert('成功匯入 ' + excelData.length + ' 筆資料！\n因為這是雛型 (Protoype)，為確保完整邏輯轉換，已在主控台印出資料結構。');
            console.log(excelData);
            closeUpload();
            // Optional: Map excelData to rawData here if format matches
        } else {
            alert('檔案似乎是空的。');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Run init
window.onload = initApp;
</script>
</body>
</html>
"""

# Write HTML content sequentially
with io.open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    f.write(HTML_HEAD_CSS)
    f.write(HTML_BODY)
    f.write(HTML_JS)

print("Dashboard V2 build successful. File located at:", OUTPUT_FILE)
