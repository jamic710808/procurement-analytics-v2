"""
採購分析儀表板 V3.0 完整版 - Vizro 風格
包含所有 10 頁的完整圖表渲染
"""

import json
import random

# 載入數據
with open('procurement_data_v3.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

company = data['company']
pages = data['pages']

# 提取各頁面數據
summary = pages['summary']
spend = pages['spend']
supplier = pages['supplier']
cost = pages['cost']
risk = pages['risk']
esg = pages['esg']
quality = pages['quality']
contract = pages['contract']
maturity = pages['maturity']
simulation = pages['simulation']

html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>採購智能分析平台 V3.0 | {company['name']}</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/js/all.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Segoe UI', 'Microsoft JhengHei', Arial, sans-serif; background: #0f172a; color: #f8fafc; overflow-x: hidden; }}

        /* Header */
        .header {{
            background: linear-gradient(135deg, #1e293b 0%, #334155 50%, #1e293b 100%);
            padding: 25px 40px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            position: relative;
            overflow: hidden;
        }}
        .header::before {{
            content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="40" fill="none" stroke="rgba(56,189,248,0.1)" stroke-width="0.5"/></svg>');
            background-size: 200px;
            opacity: 0.5;
        }}
        .header h1 {{ font-size: 28px; color: #38bdf8; margin-bottom: 8px; position: relative; }}
        .header .subtitle {{ font-size: 14px; color: #94a3b8; position: relative; }}
        .header .last-update {{ font-size: 12px; color: #64748b; margin-top: 8px; }}

        /* Navigation */
        .nav {{
            background: linear-gradient(180deg, #1e293b 0%, #0f172a 100%);
            padding: 15px 20px;
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            justify-content: center;
            border-bottom: 1px solid rgba(255,255,255,0.05);
        }}
        .nav a {{
            color: #94a3b8;
            text-decoration: none;
            padding: 10px 18px;
            border-radius: 8px;
            transition: all 0.3s ease;
            font-size: 13px;
            border: 1px solid transparent;
        }}
        .nav a:hover {{
            background: rgba(59, 130, 246, 0.2);
            color: #60a5fa;
            transform: translateY(-2px);
        }}
        .nav a.active {{
            background: linear-gradient(135deg, #3b82f6, #1d4ed8);
            color: white;
            box-shadow: 0 4px 15px rgba(59, 130, 246, 0.4);
        }}

        /* Pages */
        .page {{ display: none; padding: 30px; max-width: 1800px; margin: 0 auto; animation: fadeIn 0.4s ease; }}
        .page.active {{ display: block; }}
        @keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(20px); }} to {{ opacity: 1; transform: translateY(0); }} }}

        /* Cards */
        .card {{
            background: linear-gradient(145deg, #1e293b 0%, #172033 100%);
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 24px;
            border: 1px solid rgba(255,255,255,0.05);
            box-shadow: 0 4px 20px rgba(0,0,0,0.3);
        }}
        .card h2 {{ color: #38bdf8; font-size: 20px; margin-bottom: 20px; border-bottom: 2px solid #3b82f6; padding-bottom: 12px; display: flex; align-items: center; gap: 10px; }}
        .card h3 {{ color: #94a3b8; font-size: 15px; margin: 20px 0 15px; }}

        /* KPI Grid */
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin: 20px 0; }}
        .kpi-box {{
            background: linear-gradient(145deg, #334155, #1e293b);
            padding: 20px;
            border-radius: 12px;
            text-align: center;
            border: 1px solid rgba(255,255,255,0.05);
            transition: transform 0.3s ease;
        }}
        .kpi-box:hover {{ transform: translateY(-5px); box-shadow: 0 8px 25px rgba(0,0,0,0.3); }}
        .kpi-box.primary {{ background: linear-gradient(145deg, #3b82f6, #1d4ed8); }}
        .kpi-box.success {{ background: linear-gradient(145deg, #10b981, #059669); }}
        .kpi-box.warning {{ background: linear-gradient(145deg, #f59e0b, #d97706); }}
        .kpi-box.danger {{ background: linear-gradient(145deg, #ef4444, #dc2626); }}
        .kpi-box.info {{ background: linear-gradient(145deg, #06b6d4, #0891b2); }}
        .kpi-value {{ font-size: 32px; font-weight: 800; }}
        .kpi-label {{ font-size: 13px; opacity: 0.9; margin-top: 8px; }}
        .kpi-delta {{ font-size: 12px; margin-top: 8px; display: flex; align-items: center; justify-content: center; gap: 5px; }}
        .kpi-delta.up {{ color: #10b981; }}
        .kpi-delta.down {{ color: #ef4444; }}

        /* Charts */
        .chart-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(450px, 1fr)); gap: 24px; margin: 20px 0; }}
        .chart-container {{
            background: linear-gradient(145deg, #1e293b, #172033);
            border-radius: 12px;
            padding: 20px;
            min-height: 350px;
            border: 1px solid rgba(255,255,255,0.05);
        }}

        /* Tables */
        table {{ width: 100%; border-collapse: separate; border-spacing: 0; margin: 15px 0; font-size: 13px; }}
        th {{
            background: linear-gradient(145deg, #334155, #1e293b);
            color: #94a3b8;
            padding: 14px 16px;
            text-align: left;
            border-bottom: 2px solid #3b82f6;
            font-weight: 600;
        }}
        td {{ padding: 12px 16px; border-bottom: 1px solid #334155; }}
        tr:hover {{ background: linear-gradient(90deg, rgba(59,130,246,0.1), transparent); }}
        tr:last-child td {{ border-bottom: none; }}

        /* Badges */
        .badge {{ padding: 4px 12px; border-radius: 20px; font-size: 11px; font-weight: 600; display: inline-block; }}
        .badge-high, .badge-D {{ background: rgba(239, 68, 68, 0.2); color: #ef4444; }}
        .badge-medium, .badge-C {{ background: rgba(245, 158, 11, 0.2); color: #f59e0b; }}
        .badge-low, .badge-A {{ background: rgba(16, 185, 129, 0.2); color: #10b981; }}
        .badge-B {{ background: rgba(59, 130, 246, 0.2); color: #3b82f6; }}
        .badge-進行中 {{ background: rgba(59, 130, 246, 0.2); color: #3b82f6; }}
        .badge-規劃中 {{ background: rgba(245, 158, 11, 0.2); color: #f59e0b; }}
        .badge-已完成 {{ background: rgba(16, 185, 129, 0.2); color: #10b981; }}
        .badge-有效 {{ background: rgba(16, 185, 129, 0.2); color: #10b981; }}
        .badge-到期 {{ background: rgba(239, 68, 68, 0.2); color: #ef4444; }}

        /* Alerts */
        .alert {{ padding: 14px 18px; border-radius: 10px; margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; }}
        .alert-critical {{ background: linear-gradient(90deg, rgba(239, 68, 68, 0.2), transparent); border-left: 4px solid #ef4444; }}
        .alert-warning {{ background: linear-gradient(90deg, rgba(245, 158, 11, 0.2), transparent); border-left: 4px solid #f59e0b; }}
        .alert-info {{ background: linear-gradient(90deg, rgba(59, 130, 246, 0.2), transparent); border-left: 4px solid #3b82f6; }}

        /* Import Section */
        .import-section {{ background: linear-gradient(145deg, #1e293b, #172033); padding: 20px; border-radius: 12px; margin: 20px auto; max-width: 1800px; display: flex; justify-content: center; align-items: center; gap: 15px; border: 1px solid rgba(255,255,255,0.05); }}
        .import-section input {{ display: none; }}
        .import-section label {{ background: linear-gradient(145deg, #3b82f6, #1d4ed8); color: white; padding: 12px 24px; border-radius: 8px; cursor: pointer; transition: all 0.3s; display: flex; align-items: center; gap: 8px; }}
        .import-section label:hover {{ transform: scale(1.05); box-shadow: 0 4px 15px rgba(59,130,246,0.4); }}

        /* Progress Bars */
        .progress-bar {{ height: 8px; background: #334155; border-radius: 4px; overflow: hidden; margin-top: 5px; }}
        .progress-fill {{ height: 100%; border-radius: 4px; transition: width 1s ease; }}

        /* Responsive */
        @media (max-width: 768px) {{
            .chart-grid {{ grid-template-columns: 1fr; }}
            .kpi-grid {{ grid-template-columns: repeat(2, 1fr); }}
            .nav {{ padding: 10px; }}
            .nav a {{ padding: 8px 12px; font-size: 12px; }}
        }}
    </style>
</head>
<body>
    <!-- Header -->
    <div class="header">
        <h1><i class="fas fa-chart-line"></i> 採購智能分析平台 V3.0</h1>
        <div class="subtitle">{company['name']} | {company['industry']} | {company['reportDate']}</div>
        <div class="last-update"><i class="fas fa-clock"></i> 數據更新時間: 2025-12-15 14:30</div>
    </div>

    <!-- Navigation -->
    <nav class="nav">
        <a href="#summary" class="active" onclick="showPage('summary')"><i class="fas fa-gauge-high"></i> 決策總覽</a>
        <a href="#spend" onclick="showPage('spend')"><i class="fas fa-coins"></i> 支出分析</a>
        <a href="#supplier" onclick="showPage('supplier')"><i class="fas fa-handshake"></i> 供應商矩陣</a>
        <a href="#cost" onclick="showPage('cost')"><i class="fas fa-piggy-bank"></i> 成本節約</a>
        <a href="#risk" onclick="showPage('risk')"><i class="fas fa-shield-halved"></i> 風險管控</a>
        <a href="#esg" onclick="showPage('esg')"><i class="fas fa-leaf"></i> ESG 永續</a>
        <a href="#quality" onclick="showPage('quality')"><i class="fas fa-clipboard-check"></i> 品質 SPC</a>
        <a href="#contract" onclick="showPage('contract')"><i class="fas fa-file-contract"></i> 合約管理</a>
        <a href="#maturity" onclick="showPage('maturity')"><i class="fas fa-graduation-cap"></i> 採購成熟度</a>
        <a href="#simulation" onclick="showPage('simulation')"><i class="fas fa-flask"></i> 情境模擬</a>
    </nav>

    <!-- Import Section -->
    <div class="import-section">
        <label for="fileInput"><i class="fas fa-file-import"></i> 匯入新的 JSON 數據</label>
        <input type="file" id="fileInput" accept=".json" onchange="importData(this)">
    </div>

    <!-- Page 1: 決策總覽 -->
    <div id="summary" class="page active">
        <div class="card">
            <h2><i class="fas fa-gauge-high"></i> 關鍵績效指標 (KPI)</h2>
            <div class="kpi-grid">
                <div class="kpi-box primary">
                    <div class="kpi-value">$12.5M</div>
                    <div class="kpi-label">總採購金額</div>
                    <div class="kpi-delta up"><i class="fas fa-arrow-up"></i> +8.5% vs 去年同期</div>
                </div>
                <div class="kpi-box success">
                    <div class="kpi-value">$1.85M</div>
                    <div class="kpi-label">成本節約金額</div>
                    <div class="kpi-delta up"><i class="fas fa-arrow-up"></i> +12.3% 超額達成</div>
                </div>
                <div class="kpi-box info">
                    <div class="kpi-value">156</div>
                    <div class="kpi-label">供應商家數</div>
                    <div class="kpi-delta down"><i class="fas fa-arrow-down"></i> -3 家</div>
                </div>
                <div class="kpi-box success">
                    <div class="kpi-value">94.2%</div>
                    <div class="kpi-label">準時交貨率</div>
                    <div class="kpi-delta up"><i class="fas fa-arrow-up"></i> +2.1%</div>
                </div>
                <div class="kpi-box success">
                    <div class="kpi-value">98.7%</div>
                    <div class="kpi-label">品質合格率</div>
                    <div class="kpi-delta up"><i class="fas fa-arrow-up"></i> +0.5%</div>
                </div>
                <div class="kpi-box warning">
                    <div class="kpi-value">$1.25M</div>
                    <div class="kpi-label">風險敞口</div>
                    <div class="kpi-delta down"><i class="fas fa-arrow-down"></i> -15.2% 改善</div>
                </div>
            </div>
        </div>

        <div class="card">
            <h2><i class="fas fa-triangle-exclamation"></i> 警示與提醒</h2>
            <div class="alert alert-critical">
                <span><i class="fas fa-exclamation-circle"></i> 供應商 A 風險評級升至 HIGH - 立即檢視</span>
                <span class="badge badge-high">緊急</span>
            </div>
            <div class="alert alert-warning">
                <span><i class="fas fa-exclamation-triangle"></i> 3 項物料庫存低於安全水位 - 請儘速補貨</span>
                <span class="badge badge-medium">警告</span>
            </div>
            <div class="alert alert-info">
                <span><i class="fas fa-info-circle"></i> Q4 成本節約目標達成率 92% - 預計可超額完成</span>
                <span class="badge badge-low">資訊</span>
            </div>
        </div>

        <div class="chart-grid">
            <div class="chart-container"><div id="chart-summary-bar"></div></div>
            <div class="chart-container"><div id="chart-summary-trend"></div></div>
        </div>
    </div>

    <!-- Page 2: 支出分析 -->
    <div id="spend" class="page">
        <div class="card">
            <h2><i class="fas fa-coins"></i> 支出分析</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-spend-category"></div></div>
                <div class="chart-container"><div id="chart-spend-region"></div></div>
                <div class="chart-container"><div id="chart-spend-supplier"></div></div>
                <div class="chart-container"><div id="chart-spend-coverage"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 3: 供應商矩陣 -->
    <div id="supplier" class="page">
        <div class="card">
            <h2><i class="fas fa-handshake"></i> 供應商績效矩陣</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-supplier-scatter"></div></div>
                <div class="chart-container"><div id="chart-supplier-performance"></div></div>
                <div class="chart-container"><div id="chart-supplier-risk"></div></div>
                <div class="chart-container"><div id="chart-supplier-rating"></div></div>
            </div>
            <h3><i class="fas fa-list"></i> 供應商詳細資料</h3>
            <table>
                <tr><th>供應商</th><th>類別</th><th>採購金額</th><th>準時率</th><th>合格率</th><th>風險</th><th>評級</th></tr>
                <tr><td>供應商 A</td><td>原物料</td><td>$3,200,000</td><td>96.5%</td><td>99.2%</td><td><span class="badge badge-low">低</span></td><td><span class="badge badge-A">A</span></td></tr>
                <tr><td>供應商 B</td><td>包裝</td><td>$1,800,000</td><td>94.2%</td><td>98.5%</td><td><span class="badge badge-low">低</span></td><td><span class="badge badge-B">B</span></td></tr>
                <tr><td>供應商 C</td><td>零件</td><td>$1,500,000</td><td>89.1%</td><td>97.8%</td><td><span class="badge badge-medium">中</span></td><td><span class="badge badge-C">C</span></td></tr>
                <tr><td>供應商 D</td><td>原物料</td><td>$1,200,000</td><td>82.5%</td><td>96.2%</td><td><span class="badge badge-high">高</span></td><td><span class="badge badge-D">D</span></td></tr>
            </table>
        </div>
    </div>

    <!-- Page 4: 成本節約 -->
    <div id="cost" class="page">
        <div class="card">
            <h2><i class="fas fa-piggy-bank"></i> 成本節約分析</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-cost-type"></div></div>
                <div class="chart-container"><div id="chart-cost-target"></div></div>
                <div class="chart-container"><div id="chart-cost-variance"></div></div>
                <div class="chart-container"><div id="chart-cost-trend"></div></div>
            </div>
            <h3><i class="fas fa-history"></i> 價格差異追蹤</h3>
            <table>
                <tr><th>原物料</th><th>季度</th><th>去年單價</th><th>今年單價</th><th>差異%</th><th>狀態</th></tr>
                <tr><td>鋁合金板材</td><td>Q1</td><td>$125</td><td>$118</td><td class="kpi-delta up">-5.6%</td><td><span class="badge badge-low">節省</span></td></tr>
                <tr><td>鋁合金板材</td><td>Q2</td><td>$125</td><td>$122</td><td class="kpi-delta up">-2.4%</td><td><span class="badge badge-low">節省</span></td></tr>
                <tr><td>鋁合金板材</td><td>Q3</td><td>$125</td><td>$128</td><td class="kpi-delta down">+2.4%</td><td><span class="badge badge-medium">上漲</span></td></tr>
                <tr><td>鋁合金板材</td><td>Q4</td><td>$125</td><td>$132</td><td class="kpi-delta down">+5.6%</td><td><span class="badge badge-high">大漲</span></td></tr>
            </table>
        </div>
    </div>

    <!-- Page 5: 風險管控 -->
    <div id="risk" class="page">
        <div class="card">
            <h2><i class="fas fa-shield-halved"></i> 風險管控</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-risk-heatmap"></div></div>
                <div class="chart-container"><div id="chart-risk-trend"></div></div>
            </div>
            <h3><i class="fas fa-list"></i> 風險項目清單</h3>
            <table>
                <tr><th>ID</th><th>類別</th><th>描述</th><th>嚴重性</th><th>可能性</th><th>影響金額</th><th>狀態</th><th>負責人</th></tr>
                <tr><td>R001</td><td>供應商</td><td>單一供應商依賴過高</td><td><span class="badge badge-high">HIGH</span></td><td>可能</td><td>$250,000</td><td><span class="badge badge-進行中">進行中</span></td><td>王經理</td></tr>
                <tr><td>R002</td><td>地緣政治</td><td>亞洲供應鏈風險</td><td><span class="badge badge-medium">MEDIUM</span></td><td>可能</td><td>$180,000</td><td><span class="badge badge-規劃中">規劃中</span></td><td>李經理</td></tr>
                <tr><td>R003</td><td>品質</td><td>關鍵物料合格率下降</td><td><span class="badge badge-high">HIGH</span></td><td>偶爾</td><td>$320,000</td><td><span class="badge badge-已完成">已完成</span></td><td>陳經理</td></tr>
            </table>
        </div>
    </div>

    <!-- Page 6: ESG 永續 -->
    <div id="esg" class="page">
        <div class="card">
            <h2><i class="fas fa-leaf"></i> ESG 永續評核</h2>
            <div class="kpi-grid">
                <div class="kpi-box success">
                    <div class="kpi-value">72</div>
                    <div class="kpi-label">環境 (E) 評分</div>
                    <div class="progress-bar"><div class="progress-fill" style="width: 72%; background: #10b981;"></div></div>
                </div>
                <div class="kpi-box warning">
                    <div class="kpi-value">68</div>
                    <div class="kpi-label">社會 (S) 評分</div>
                    <div class="progress-bar"><div class="progress-fill" style="width: 68%; background: #f59e0b;"></div></div>
                </div>
                <div class="kpi-box info">
                    <div class="kpi-value">78</div>
                    <div class="kpi-label">治理 (G) 評分</div>
                    <div class="progress-bar"><div class="progress-fill" style="width: 78%; background: #06b6d4;"></div></div>
                </div>
            </div>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-esg-radar"></div></div>
                <div class="chart-container"><div id="chart-esg-carbon"></div></div>
                <div class="chart-container"><div id="chart-esg-supplier"></div></div>
                <div class="chart-container"><div id="chart-esg-goals"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 7: 品質 SPC -->
    <div id="quality" class="page">
        <div class="card">
            <h2><i class="fas fa-clipboard-check"></i> 品質 SPC 管制</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-quality-cpk"></div></div>
                <div class="chart-container"><div id="chart-quality-defect"></div></div>
                <div class="chart-container"><div id="chart-quality-pareto"></div></div>
                <div class="chart-container"><div id="chart-quality-supplier"></div></div>
            </div>
            <h3><i class="fas fa-industry"></i> 供應商品質排名</h3>
            <table>
                <tr><th>供應商</th><th>不良 PPM</th><th>品質評分</th><th>準時率</th><th>回應時間</th><th>整體評級</th></tr>
                <tr><td>供應商 A</td><td>850</td><td>98.5</td><td>96.5%</td><td>2.5 天</td><td><span class="badge badge-A">A</span></td></tr>
                <tr><td>供應商 B</td><td>1,200</td><td>96.2</td><td>94.2%</td><td>4.2 天</td><td><span class="badge badge-B">B</span></td></tr>
                <tr><td>供應商 C</td><td>2,100</td><td>92.8</td><td>89.1%</td><td>8.5 天</td><td><span class="badge badge-C">C</span></td></tr>
                <tr><td>供應商 D</td><td>3,500</td><td>88.5</td><td>82.5%</td><td>12.0 天</td><td><span class="badge badge-D">D</span></td></tr>
            </table>
        </div>
    </div>

    <!-- Page 8: 合約管理 -->
    <div id="contract" class="page">
        <div class="card">
            <h2><i class="fas fa-file-contract"></i> 合約生命週期管理</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-contract-value"></div></div>
                <div class="chart-container"><div id="chart-contract-timeline"></div></div>
            </div>
            <h3><i class="fas fa-list"></i> 合約清單</h3>
            <table>
                <tr><th>合約編號</th><th>供應商</th><th>類型</th><th>金額</th><th>起始日</th><th>到期日</th><th>狀態</th><th>續約提醒</th></tr>
                <tr><td>C001</td><td>供應商 A</td><td>年度框架</td><td>$2,500,000</td><td>2025-01-01</td><td>2025-12-31</td><td><span class="badge badge-有效">有效</span></td><td>Q4</td></tr>
                <tr><td>C002</td><td>供應商 B</td><td>2年期</td><td>$1,800,000</td><td>2024-07-01</td><td>2026-06-30</td><td><span class="badge badge-有效">有效</span></td><td>Q2</td></tr>
                <tr><td>C003</td><td>供應商 C</td><td>單次採購</td><td>$850,000</td><td>2025-03-01</td><td>2025-08-31</td><td><span class="badge badge-到期">到期</span></td><td>-</td></tr>
            </table>
            <h3><i class="fas fa-exclamation-circle"></i> 即將到期合約</h3>
            <table>
                <tr><th>合約</th><th>供應商</th><th>剩餘天數</th><th>金額</th><th>建議動作</th></tr>
                <tr><td>C003</td><td>供應商 C</td><td class="kpi-delta down">15 天</td><td>$850,000</td><td><span class="badge badge-high">立即續約</span></td></tr>
                <tr><td>C006</td><td>供應商 F</td><td class="kpi-delta down">45 天</td><td>$620,000</td><td><span class="badge badge-medium">評估談判</span></td></tr>
            </table>
        </div>
    </div>

    <!-- Page 9: 採購成熟度 -->
    <div id="maturity" class="page">
        <div class="card">
            <h2><i class="fas fa-graduation-cap"></i> 採購成熟度評估</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-maturity-radar"></div></div>
                <div class="chart-container"><div id="chart-maturity-current"></div></div>
                <div class="chart-container"><div id="chart-maturity-roadmap"></div></div>
                <div class="chart-container"><div id="chart-maturity-benchmark"></div></div>
            </div>
            <h3><i class="fas fa-list"></i> 各維度評分</h3>
            <table>
                <tr><th>維度</th><th>分數</th><th>目標</th><th>進度</th><th>說明</th></tr>
                <tr><td>策略採購</td><td><span class="badge badge-B">3.2/5</span></td><td>4.0</td><td><div class="progress-bar"><div class="progress-fill" style="width: 64%; background: #3b82f6;"></div></div></td><td>建立採購策略框架</td></tr>
                <tr><td>供應商管理</td><td><span class="badge badge-A">3.8/5</span></td><td>4.2</td><td><div class="progress-bar"><div class="progress-fill" style="width: 76%; background: #10b981;"></div></div></td><td>完整供應商評估制度</td></tr>
                <tr><td>流程優化</td><td><span class="badge badge-C">2.9/5</span></td><td>3.5</td><td><div class="progress-bar"><div class="progress-fill" style="width: 58%; background: #f59e0b;"></div></div></td><td>標準化流程進行中</td></tr>
                <tr><td>數據分析</td><td><span class="badge badge-C">2.5/5</span></td><td>3.8</td><td><div class="progress-bar"><div class="progress-fill" style="width: 50%; background: #f59e0b;"></div></div></td><td>基礎數據收集階段</td></tr>
            </table>
        </div>
    </div>

    <!-- Page 10: 情境模擬 -->
    <div id="simulation" class="page">
        <div class="card">
            <h2><i class="fas fa-flask"></i> 情境模擬站</h2>

            <div class="card" style="background: linear-gradient(145deg, #334155, #1e293b);">
                <h3><i class="fas fa-bolt"></i> 情境 1: 供應商斷鏈</h3>
                <p style="color: #94a3b8; margin: 10px 0;">假設主要供應商因天災停止供貨 30 天</p>
                <div class="kpi-grid">
                    <div class="kpi-box danger"><div class="kpi-label">生產延誤</div><div class="kpi-value" style="font-size:20px">15 天</div></div>
                    <div class="kpi-box danger"><div class="kpi-label">預估損失</div><div class="kpi-value" style="font-size:20px">$450,000</div></div>
                    <div class="kpi-box warning"><div class="kpi-label">替代來源</div><div class="kpi-value" style="font-size:20px">7 天切換</div></div>
                    <div class="kpi-box success"><div class="kpi-label">緩解成本</div><div class="kpi-value" style="font-size:20px">$125,000</div></div>
                </div>
            </div>

            <div class="card" style="background: linear-gradient(145deg, #334155, #1e293b);">
                <h3><i class="fas fa-chart-line"></i> 情境 2: 原物料漲價 20%</h3>
                <p style="color: #94a3b8; margin: 10px 0;">鋁合金市場價格突然飆漲</p>
                <div class="kpi-grid">
                    <div class="kpi-box danger"><div class="kpi-label">年度成本增加</div><div class="kpi-value" style="font-size:20px">$1,360,000</div></div>
                    <div class="kpi-box warning"><div class="kpi-label">毛利影響</div><div class="kpi-value" style="font-size:20px">-2.5%</div></div>
                    <div class="kpi-box info"><div class="kpi-label">價格轉嫁</div><div class="kpi-value" style="font-size:20px">60%</div></div>
                </div>
            </div>

            <div class="chart-grid">
                <div class="chart-container"><div id="chart-sim-heatmap"></div></div>
                <div class="chart-container"><div id="chart-sim-sensitivity"></div></div>
            </div>
        </div>
    </div>

    <script>
    // 全域數據
    const dashboardData = {json.dumps(data, ensure_ascii=False)};

    // 導航函數
    function showPage(pageId) {{
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.nav a').forEach(a => a.classList.remove('active'));
        document.getElementById(pageId).classList.add('active');
        document.querySelector('a[href="#' + pageId + '"]').classList.add('active');

        // 渲染該頁面的圖表
        setTimeout(() => renderCharts(pageId), 100);
    }}

    // 圖表渲染函數
    function renderCharts(pageId) {{
        const layoutDark = {{title: {{font: {{color: '#94a3b8'}}, x: 0.05}}, paper_bgcolor: 'transparent', plot_bgcolor: 'rgba(0,0,0,0)', font: {{color: '#94a3b8'}}, margin: {{t: 50, l: 50, r: 50, b: 50}}}};

        // Page 1: 決策總覽
        if (pageId === 'summary') {{
            // KPI 分布圖
            Plotly.newPlot('chart-summary-bar', [{{
                x: ['原物料', '包裝', '設備', '物流', '其他'],
                y: [54.4, 16.8, 14.8, 8.8, 5.2],
                type: 'bar', marker: {{color: ['#3b82f6', '#10b981', '#f59e0b', '#06b6d4', '#8b5cf6']}},
                text: '54.4%', textposition: 'outside'
            }}], {{...layoutDark, title: '支出分布'}});

            // 趨勢圖
            Plotly.newPlot('chart-summary-trend', [{{
                x: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
                y: [1.8, 2.1, 1.9, 2.3, 2.0, 2.4],
                type: 'scatter', mode: 'lines+markers',
                line: {{color: '#10b981', width: 3}},
                fill: 'tozeroy', fillcolor: 'rgba(16, 185, 129, 0.2)'
            }}], {{...layoutDark, title: '成本節約趨勢 (百萬美元)'}});
        }}

        // Page 2: 支出分析
        if (pageId === 'spend') {{
            // 類別分布
            Plotly.newPlot('chart-spend-category', [{{
                labels: ['原物料', '包裝', '設備', '物流', '其他'],
                values: [54.4, 16.8, 14.8, 8.8, 5.2],
                type: 'pie', hole: 0.4,
                marker: {{colors: ['#3b82f6', '#10b981', '#f59e0b', '#06b6d4', '#8b5cf6']}}
            }}], {{...layoutDark, title: '支出類別分布'}});

            // 區域分布
            Plotly.newPlot('chart-spend-region', [{{
                labels: ['亞太', '美洲', '歐洲'],
                values: [60, 20, 20],
                type: 'pie', hole: 0.4,
                marker: {{colors: ['#3b82f6', '#10b981', '#f59e0b']}}
            }}], {{...layoutDark, title: '區域支出分布'}});

            // 供應商分布
            Plotly.newPlot('chart-spend-supplier', [{{
                x: ['供應商A', '供應商B', '供應商C', '供應商D', '其他'],
                y: [3200, 1800, 1500, 1200, 4800],
                type: 'bar', marker: {{color: '#3b82f6'}}
            }}], {{...layoutDark, title: '供應商採購金額 (千美元)'}});

            // 合約覆蓋率
            Plotly.newPlot('chart-spend-coverage', [{{
                x: ['原物料', '包裝', '設備', '物流', '其他'],
                y: [95, 88, 72, 65, 35],
                type: 'bar', marker: {{color: '#10b981'}},
                text: '95%', textposition: 'outside'
            }}], {{...layoutDark, title: '合約覆蓋率 (%)'}});
        }}

        // Page 3: 供應商矩陣
        if (pageId === 'supplier') {{
            // 散點圖
            Plotly.newPlot('chart-supplier-scatter', [{{
                x: [92, 88, 85, 78, 95, 72, 82, 90],
                y: [15, 25, 45, 60, 10, 75, 35, 20],
                mode: 'markers+text',
                marker: {{size: 20, color: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444', '#10b981', '#ef4444', '#3b82f6', '#10b981']}},
                text: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'], textposition: 'top center'
            }}], {{...layoutDark, title: '供應商績效 vs 風險'}});

            // 供應商評分
            Plotly.newPlot('chart-supplier-performance', [{{
                x: ['供應商A', '供應商B', '供應商C', '供應商D'],
                y: [[96.5, 94.2, 89.1, 82.5]],
                type: 'bar', name: '準時率'
            }}, {{
                x: ['供應商A', '供應商B', '供應商C', '供應商D'],
                y: [[99.2, 98.5, 97.8, 96.2]],
                type: 'bar', name: '合格率'
            }}], {{...layoutDark, barmode: 'group', title: '供應商表現對比'}});

            // 風險分布
            Plotly.newPlot('chart-supplier-risk', [{{
                labels: ['低風險', '中風險', '高風險'],
                values: [45, 78, 33],
                type: 'pie', hole: 0.4,
                marker: {{colors: ['#10b981', '#f59e0b', '#ef4444']}}
            }}], {{...layoutDark, title: '供應商風險分布'}});

            // 評級分布
            Plotly.newPlot('chart-supplier-rating', [{{
                x: ['A', 'B', 'C', 'D'],
                y: [3, 2, 2, 1],
                type: 'bar', marker: {{color: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444']}},
                text: '3', textposition: 'outside'
            }}], {{...layoutDark, title: '供應商評級分布'}});
        }}

        // Page 4: 成本節約
        if (pageId === 'cost') {{
            // 節約類型
            Plotly.newPlot('chart-cost-type', [{{
                labels: ['價格談判', '規格優化', '數量折扣', '供應商整合', '替代材料'],
                values: [44.3, 24.3, 17.3, 9.7, 4.3],
                type: 'pie', hole: 0.4
            }}], {{...layoutDark, title: '成本節約類型分布'}});

            // 目標達成
            Plotly.newPlot('chart-cost-target', [{{
                x: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
                y: [150, 150, 160, 160, 165, 165],
                type: 'scatter', mode: 'lines', name: '目標',
                line: {{color: '#ef4444', dash: 'dash'}}
            }}, {{
                x: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
                y: [145, 162, 158, 175, 168, 182],
                type: 'scatter', mode: 'lines+markers', name: '實際',
                fill: 'tonexty', fillcolor: 'rgba(16, 185, 129, 0.2)',
                line: {{color: '#10b981'}}
            }}], {{...layoutDark, title: '成本節約目標 vs 實際'}});

            // 價格差異
            Plotly.newPlot('chart-cost-variance', [{{
                x: ['Q1', 'Q2', 'Q3', 'Q4'],
                y: [-5.6, -2.4, 2.4, 5.6],
                type: 'bar', marker: {{color: ['#10b981', '#10b981', '#ef4444', '#ef4444']}},
                text: ['-5.6%', '-2.4%', '+2.4%', '+5.6%'], textposition: 'outside'
            }}], {{...layoutDark, title: '價格差異 (%)'}});

            // 趨勢
            Plotly.newPlot('chart-cost-trend', [{{
                x: ['2022', '2023', '2024', '2025'],
                y: [1.2, 1.45, 1.68, 1.85],
                type: 'scatter', mode: 'lines+markers',
                line: {{color: '#3b82f6', width: 3}},
                fill: 'tozeroy', fillcolor: 'rgba(59, 130, 246, 0.2)'
            }}], {{...layoutDark, title: '年度成本節約趨勢 (百萬美元)'}});
        }}

        // Page 5: 風險管控
        if (pageId === 'risk') {{
            // 熱力圖
            Plotly.newPlot('chart-risk-heatmap', [{{
                x: ['RARE', 'OCCASIONAL', 'POSSIBLE', 'LIKELY'],
                y: ['輕微', '嚴重'],
                z: [[8, 1], [3, 2]],
                type: 'heatmap', colorscale: [[0, '#10b981'], [0.5, '#f59e0b'], [1, '#ef4444']]
            }}], {{...layoutDark, title: '風險熱力圖'}});

            // 趨勢
            Plotly.newPlot('chart-risk-trend', [{{
                x: ['Q1', 'Q2', 'Q3', 'Q4'],
                y: [5, 4, 3, 2],
                type: 'scatter', mode: 'lines+markers',
                line: {{color: '#ef4444', width: 3}}
            }}], {{...layoutDark, title: '風險項目數趨勢'}});
        }}

        // Page 6: ESG
        if (pageId === 'esg') {{
            // 雷達圖
            Plotly.newPlot('chart-esg-radar', [{{
                r: [72, 68, 78, 65, 70],
                theta: ['E', 'S', 'G', '供應商', '減碳'],
                fill: 'toself', marker: {{color: '#10b981'}}
            }}], {{...layoutDark, title: 'ESG 雷達圖', polar: {{radialaxis: {{visible: true, range: [0, 100]}}}}}});

            // 碳排放
            Plotly.newPlot('chart-esg-carbon', [{{
                x: ['2022', '2023', '2024', '2025'],
                y: [12500, 11800, 10500, 9200],
                type: 'bar', marker: {{color: '#10b981'}},
                text: '12,500', textposition: 'outside'
            }}], {{...layoutDark, title: '碳排放趨勢 (噸 CO2e)'}});

            // 供應商 ESG
            Plotly.newPlot('chart-esg-supplier', [{{
                x: ['供應商A', '供應商B', '供應商C', '供應商D'],
                y: [85, 72, 65, 55],
                type: 'bar', name: '環境'
            }}], {{...layoutDark, title: '供應商 ESG 評分'}});

            // 目標
            Plotly.newPlot('chart-esg-goals', [{{
                x: ['E', 'S', 'G'],
                y: [[72, 80], [68, 75], [78, 80]],
                type: 'bar', barmode: 'group'
            }}], {{...layoutDark, title: 'ESG 目標 vs 實際'}});
        }}

        // Page 7: 品質
        if (pageId === 'quality') {{
            // CPK
            Plotly.newPlot('chart-quality-cpk', [{{
                x: ['W1', 'W2', 'W3', 'W4', 'W5', 'W6'],
                y: [1.45, 1.52, 1.38, 1.55, 1.48, 1.62],
                type: 'scatter', mode: 'lines+markers',
                line: {{color: '#3b82f6', width: 3}},
                fill: 'tozeroy', fillcolor: 'rgba(59, 130, 246, 0.2)'
            }}, {{
                x: ['W1', 'W2', 'W3', 'W4', 'W5', 'W6'],
                y: [1.33, 1.33, 1.33, 1.33, 1.33, 1.33],
                type: 'scatter', mode: 'lines',
                line: {{color: '#ef4444', dash: 'dash'}}
            }}], {{...layoutDark, title: 'CPK 趨勢'}});

            // 異常類型
            Plotly.newPlot('chart-quality-defect', [{{
                x: ['尺寸', '外觀', '功能', '包裝', '標籤'],
                y: [45, 38, 28, 18, 11],
                type: 'bar', marker: {{color: '#ef4444'}}
            }}], {{...layoutDark, title: '品質異常分布'}});

            // 柏拉圖
            Plotly.newPlot('chart-quality-pareto', [{{
                x: ['尺寸偏差', '外觀瑕疵', '功能異常', '包裝損壞', '標籤錯誤'],
                y: [45, 38, 28, 18, 11],
                type: 'bar', name: '次數'
            }}], {{...layoutDark, title: '品質異常柏拉圖'}});

            // 供應商品質
            Plotly.newPlot('chart-quality-supplier', [{{
                x: ['A', 'B', 'C', 'D'],
                y: [850, 1200, 2100, 3500],
                type: 'bar', marker: {{color: '#ef4444'}}
            }}], {{...layoutDark, title: '供應商不良 PPM'}});
        }}

        // Page 8: 合約
        if (pageId === 'contract') {{
            // 價值分布
            Plotly.newPlot('chart-contract-value', [{{
                x: ['C001', 'C002', 'C003', 'C004', 'C005'],
                y: [2500, 1800, 850, 1200, 3500],
                type: 'bar', marker: {{color: '#3b82f6'}}
            }}], {{...layoutDark, title: '合約金額分布 (千美元)'}});

            // 時間軸
            Plotly.newPlot('chart-contract-timeline', [{{
                x: ['C001', 'C002', 'C003', 'C004', 'C005'],
                y: [1, 2, 0.5, 1.5, 3],
                type: 'bar', orientation: 'h'
            }}], {{...layoutDark, title: '合約期限 (年)'}});
        }}

        // Page 9: 成熟度
        if (pageId === 'maturity') {{
            // 雷達圖
            Plotly.newPlot('chart-maturity-radar', [{{
                r: [3.2, 3.8, 2.9, 2.5, 3.5, 2.8],
                theta: ['策略', '供應商', '流程', '數據', '風險', '永續'],
                fill: 'toself', marker: {{color: '#8b5cf6'}}
            }}], {{...layoutDark, polar: {{radialaxis: {{visible: true, range: [0, 5]}}}}, title: '成熟度雷達圖'}});

            // 當前評分
            Plotly.newPlot('chart-maturity-current', [{{
                x: ['策略採購', '供應商', '流程', '數據', '風險', '永續'],
                y: [3.2, 3.8, 2.9, 2.5, 3.5, 2.8],
                type: 'bar', marker: {{color: '#8b5cf6'}}
            }}], {{...layoutDark, title: '當前成熟度評分'}});

            // 藍圖
            Plotly.newPlot('chart-maturity-roadmap', [{{
                x: ['2025', '2026', '2027'],
                y: [3.2, 3.8, 4.5],
                type: 'scatter', mode: 'lines+markers',
                line: {{color: '#10b981', width: 3}},
                fill: 'tozeroy', fillcolor: 'rgba(16, 185, 129, 0.2)'
            }}], {{...layoutDark, title: '成熟度發展藍圖'}});

            // 標竿
            Plotly.newPlot('chart-maturity-benchmark', [{{
                labels: ['初始', '發展', '定義', '優化'],
                values: [15, 35, 32, 18],
                type: 'pie', hole: 0.4
            }}], {{...layoutDark, title: '產業標竿分布'}});
        }}

        // Page 10: 模擬
        if (pageId === 'simulation') {{
            // 熱力圖
            Plotly.newPlot('chart-sim-heatmap', [{{
                x: ['供應商斷鏈', '漲價20%', '匯率波動'],
                y: ['損失', '緩解成本', '效果'],
                z: [[450, 125, 85], [1360, 180, 90], [320, 45, 70]],
                type: 'heatmap'
            }}], {{...layoutDark, title: '情境分析矩陣'}});

            // 敏感度
            Plotly.newPlot('chart-sim-sensitivity', [{{
                x: [-15, -10, -5, 0, 5, 10, 15, 18],
                y: [-5, -3, -1, 0, 2, 5, 8, 12],
                type: 'scatter', mode: 'lines',
                line: {{color: '#f59e0b', width: 3}}
            }}], {{...layoutDark, title: '敏感度分析'}});
        }}
    }}

    // 數據匯入
    function importData(input) {{
        const file = input.files[0];
        const reader = new FileReader();
        reader.onload = function(e) {{
            try {{
                const newData = JSON.parse(e.target.result);
                alert('數據匯入成功！請重新生成儀表板。');
            }} catch(err) {{
                alert('JSON 格式錯誤: ' + err.message);
            }}
        }};
        reader.readAsText(file);
    }}

    // 初始渲染
    renderCharts('summary');
    </script>
</body>
</html>'''

# 寫入檔案
with open('procurement_dashboard_v3_full.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("OK Full V3.0 Dashboard generated: procurement_dashboard_v3_full.html")
print("   10 pages with complete chart rendering")
