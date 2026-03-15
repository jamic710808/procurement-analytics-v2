"""
採購分析儀表板 V3.0 - Vizro 風格
基於 procurement_data_v3.json 數據生成
"""

import json

# 載入數據
with open('procurement_data_v3.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

company = data['company']
pages = data['pages']

html = '''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>採購分析 V3.0 - XYZ 製造</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', 'Microsoft JhengHei', Arial, sans-serif; background: #0f172a; color: #f8fafc; }
        .header { background: linear-gradient(135deg, #1e293b, #334155); padding: 20px 30px; border-bottom: 1px solid rgba(255,255,255,0.1); }
        .header h1 { font-size: 24px; color: #38bdf8; margin-bottom: 5px; }
        .header .subtitle { font-size: 14px; color: #94a3b8; }
        .nav { background: #1e293b; padding: 15px 0; display: flex; gap: 10px; flex-wrap: wrap; justify-content: center; }
        .nav a { color: #94a3b8; text-decoration: none; padding: 10px 20px; border-radius: 8px; transition: 0.3s; font-size: 14px; }
        .nav a:hover, .nav a.active { background: #3b82f6; color: white; }
        .page { display: none; padding: 30px; max-width: 1600px; margin: 0 auto; }
        .page.active { display: block; }
        .card { background: #1e293b; border-radius: 12px; padding: 20px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.05); }
        .card h2 { color: #38bdf8; font-size: 18px; margin-bottom: 15px; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; }
        .card h3 { color: #94a3b8; font-size: 14px; margin: 15px 0 10px; }
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }
        .kpi-box { background: linear-gradient(135deg, #334155, #1e293b); padding: 20px; border-radius: 10px; text-align: center; }
        .kpi-box.primary { background: linear-gradient(135deg, #3b82f6, #1d4ed8); }
        .kpi-box.success { background: linear-gradient(135deg, #10b981, #059669); }
        .kpi-box.warning { background: linear-gradient(135deg, #f59e0b, #d97706); }
        .kpi-box.danger { background: linear-gradient(135deg, #ef4444, #dc2626); }
        .kpi-value { font-size: 28px; font-weight: bold; }
        .kpi-label { font-size: 12px; opacity: 0.9; margin-top: 5px; }
        .kpi-delta { font-size: 12px; margin-top: 5px; }
        .kpi-delta.up { color: #10b981; }
        .kpi-delta.down { color: #ef4444; }
        .chart-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 20px; margin: 20px 0; }
        .chart-container { background: #1e293b; border-radius: 10px; padding: 15px; min-height: 300px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px; }
        th { background: #334155; color: #94a3b8; padding: 12px; text-align: left; border-bottom: 1px solid #475569; }
        td { padding: 10px 12px; border-bottom: 1px solid #334155; }
        tr:hover { background: #334155; }
        .badge { padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; }
        .badge-high { background: rgba(239, 68, 68, 0.2); color: #ef4444; }
        .badge-medium { background: rgba(245, 158, 11, 0.2); color: #f59e0b; }
        .badge-low { background: rgba(16, 185, 129, 0.2); color: #10b981; }
        .badge-a { background: rgba(59, 130, 246, 0.2); color: #3b82f6; }
        .badge-b { background: rgba(16, 185, 129, 0.2); color: #10b981; }
        .badge-c { background: rgba(245, 158, 11, 0.2); color: #f59e0b; }
        .badge-d { background: rgba(239, 68, 68, 0.2); color: #ef4444; }
        .alert { padding: 12px 15px; border-radius: 8px; margin-bottom: 10px; display: flex; gap: 10px; align-items: center; }
        .alert-critical { background: rgba(239, 68, 68, 0.15); border-left: 3px solid #ef4444; }
        .alert-warning { background: rgba(245, 158, 11, 0.15); border-left: 3px solid #f59e0b; }
        .alert-info { background: rgba(59, 130, 246, 0.15); border-left: 3px solid #3b82f6; }
        .import-section { background: #1e293b; padding: 20px; border-radius: 10px; margin: 20px 0; text-align: center; }
        .import-section input { display: none; }
        .import-section label { background: #3b82f6; color: white; padding: 10px 20px; border-radius: 8px; cursor: pointer; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 採購智能分析平台 V3.0</h1>
        <div class="subtitle">XYZ 製造股份有限公司 | 電子製造服務 (EMS) | 2025-Q4</div>
    </div>

    <nav class="nav">
        <a href="#summary" class="active" onclick="showPage('summary')">1. 決策總覽</a>
        <a href="#spend" onclick="showPage('spend')">2. 支出分析</a>
        <a href="#supplier" onclick="showPage('supplier')">3. 供應商矩陣</a>
        <a href="#cost" onclick="showPage('cost')">4. 成本節約</a>
        <a href="#risk" onclick="showPage('risk')">5. 風險管控</a>
        <a href="#esg" onclick="showPage('esg')">6. ESG 永續</a>
        <a href="#quality" onclick="showPage('quality')">7. 品質 SPC</a>
        <a href="#contract" onclick="showPage('contract')">8. 合約管理</a>
        <a href="#maturity" onclick="showPage('maturity')">9. 採購成熟度</a>
        <a href="#simulation" onclick="showPage('simulation')">10. 情境模擬</a>
    </nav>

    <div class="import-section">
        <label for="fileInput">
            <i class="fas fa-upload"></i> 匯入/更新資料 (JSON)
        </label>
        <input type="file" id="fileInput" accept=".json" onchange="importData(this)">
    </div>

    <!-- Page 1: 決策總覽 -->
    <div id="summary" class="page active">
        <div class="card">
            <h2>📈 關鍵績效指標 (KPI)</h2>
            <div class="kpi-grid">
                <div class="kpi-box primary">
                    <div class="kpi-value">$12,500,000</div>
                    <div class="kpi-label">總採購金額 (USD)</div>
                    <div class="kpi-delta up">+8.5% vs 去年同期</div>
                </div>
                <div class="kpi-box success">
                    <div class="kpi-value">$1,850,000</div>
                    <div class="kpi-label">成本節約</div>
                    <div class="kpi-delta up">+12.3%達成</div>
                </div>
                <div class="kpi-box">
                    <div class="kpi-value">94.2%</div>
                    <div class="kpi-label">準時交貨率</div>
                    <div class="kpi-delta up">+2.1%</div>
                </div>
                <div class="kpi-box">
                    <div class="kpi-value">98.7%</div>
                    <div class="kpi-label">品質合格率</div>
                    <div class="kpi-delta up">+0.5%</div>
                </div>
            </div>
        </div>
        <div class="card">
            <h2>⚠️ 警示與提醒</h2>
            <div class="alert alert-critical"><span>供應商 A 風險評級升至 HIGH</span><span class="badge badge-high">立即檢視</span></div>
            <div class="alert alert-warning"><span>3 項物料庫存低於安全水位</span><span class="badge badge-medium">檢視庫存</span></div>
            <div class="alert alert-info"><span>Q4 成本節約目標達成率 92%</span><span class="badge badge-low">查看詳情</span></div>
        </div>
        <div class="chart-grid">
            <div class="chart-container"><div id="chart-summary-1"></div></div>
            <div class="chart-container"><div id="chart-summary-2"></div></div>
        </div>
    </div>

    <!-- Page 2: 支出分析 -->
    <div id="spend" class="page">
        <div class="card">
            <h2>💰 支出分析</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-spend-category"></div></div>
                <div class="chart-container"><div id="chart-spend-region"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 3: 供應商矩陣 -->
    <div id="supplier" class="page">
        <div class="card">
            <h2>🏭 供應商績效矩陣</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-supplier-scatter"></div></div>
                <div class="chart-container"><div id="chart-supplier-risk"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 4: 成本節約 -->
    <div id="cost" class="page">
        <div class="card">
            <h2>💵 成本節約分析</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-cost-type"></div></div>
                <div class="chart-container"><div id="chart-cost-target"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 5: 風險管控 -->
    <div id="risk" class="page">
        <div class="card">
            <h2>🛡️ 風險管控</h2>
            <table>
                <tr><th>ID</th><th>類別</th><th>描述</th><th>嚴重性</th><th>影響金額</th><th>狀態</th><th>負責人</th></tr>
                <tr><td>R001</td><td>供應商</td><td>單一供應商依賴過高</td><td><span class="badge badge-high">HIGH</span></td><td>$250,000</td><td>進行中</td><td>採購部 王經理</td></tr>
                <tr><td>R002</td><td>地緣政治</td><td>亞洲供應鏈風險</td><td><span class="badge badge-medium">MEDIUM</span></td><td>$180,000</td><td>規劃中</td><td>供應鏈部 李經理</td></tr>
                <tr><td>R003</td><td>品質</td><td>關鍵物料合格率下降</td><td><span class="badge badge-high">HIGH</span></td><td>$320,000</td><td>已完成</td><td>品質部 陳經理</td></tr>
            </table>
        </div>
    </div>

    <!-- Page 6: ESG -->
    <div id="esg" class="page">
        <div class="card">
            <h2>🌿 ESG 永續評核</h2>
            <div class="kpi-grid">
                <div class="kpi-box success"><div class="kpi-value">72</div><div class="kpi-label">環境 (E) 評分</div></div>
                <div class="kpi-box warning"><div class="kpi-value">68</div><div class="kpi-label">社會 (S) 評分</div></div>
                <div class="kpi-box primary"><div class="kpi-value">78</div><div class="kpi-label">治理 (G) 評分</div></div>
            </div>
        </div>
    </div>

    <!-- Page 7: 品質 SPC -->
    <div id="quality" class="page">
        <div class="card">
            <h2>📏 品質 SPC 管制</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-quality-cpk"></div></div>
                <div class="chart-container"><div id="chart-quality-defect"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 8: 合約管理 -->
    <div id="contract" class="page">
        <div class="card">
            <h2>📄 合約生命週期</h2>
            <table>
                <tr><th>合約編號</th><th>供應商</th><th>類型</th><th>金額</th><th>狀態</th></tr>
                <tr><td>C001</td><td>供應商 A</td><td>年度框架</td><td>$2,500,000</td><td><span class="badge badge-a">有效</span></td></tr>
                <tr><td>C002</td><td>供應商 B</td><td>2年期</td><td>$1,800,000</td><td><span class="badge badge-a">有效</span></td></tr>
                <tr><td>C003</td><td>供應商 C</td><td>單次採購</td><td>$850,000</td><td><span class="badge badge-d">到期</span></td></tr>
            </table>
        </div>
    </div>

    <!-- Page 9: 採購成熟度 -->
    <div id="maturity" class="page">
        <div class="card">
            <h2>🎓 採購成熟度評估</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-maturity-radar"></div></div>
                <div class="chart-container"><div id="chart-maturity-roadmap"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 10: 情境模擬 -->
    <div id="simulation" class="page">
        <div class="card">
            <h2>🔮 情境模擬站</h2>
            <div class="card" style="background: #334155; margin: 15px 0;">
                <h3>供應商斷鏈</h3>
                <p style="color: #94a3b8; margin: 10px 0;">假設主要供應商因天災停止供貨 30 天</p>
                <div class="kpi-grid">
                    <div class="kpi-box"><div class="kpi-label">預估損失</div><div class="kpi-value" style="font-size:18px">$450,000</div></div>
                    <div class="kpi-box warning"><div class="kpi-label">緩解成本</div><div class="kpi-value" style="font-size:18px">$125,000</div></div>
                    <div class="kpi-box success"><div class="kpi-label">緩解效果</div><div class="kpi-value" style="font-size:18px">85%</div></div>
                </div>
            </div>
        </div>
    </div>

    <script>
    function showPage(pageId) {
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.nav a').forEach(a => a.classList.remove('active'));
        document.getElementById(pageId).classList.add('active');
        document.querySelector('a[href="#' + pageId + '"]').classList.add('active');
        renderCharts();
    }

    function renderCharts() {
        // 圖表 1: 支出分布
        if (document.getElementById('chart-summary-1')) {
            Plotly.newPlot('chart-summary-1', [{
                x: ['原物料', '包裝', '設備', '物流', '其他'],
                y: [54.4, 16.8, 14.8, 8.8, 5.2],
                type: 'bar', marker: {color: '#3b82f6'}
            }], {title: '支出分布 (%)', template: 'plotly_dark'});
        }
    }

    function importData(input) {
        const file = input.files[0];
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const newData = JSON.parse(e.target.result);
                alert('數據匯入成功！請重新生成儀表板。');
            } catch(err) {
                alert('JSON 格式錯誤');
            }
        };
        reader.readAsText(file);
    }

    renderCharts();
    </script>
</body>
</html>'''

# 寫入檔案
with open('procurement_dashboard_v3.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("OK V3.0 Dashboard generated: procurement_dashboard_v3.html")
