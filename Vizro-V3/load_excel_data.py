"""
採購分析儀表板 V3.0 - 從 Excel 讀取數據生成儀表板
"""

import pandas as pd
import json

def load_data_from_excel(excel_file='procurement_data_v3.xlsx'):
    """從 Excel 讀取所有數據並轉換為儀表板格式"""

    excel_data = pd.ExcelFile(excel_file)
    sheets = excel_data.sheet_names

    # 建立數據結構
    data = {
        "company": {
            "name": "XYZ 製造股份有限公司",
            "code": "XYZMFG",
            "industry": "電子製造服務 (EMS)",
            "reportDate": "2025-Q4",
            "currency": "USD"
        },
        "pages": {}
    }

    # 讀取每個工作表
    for sheet_name in sheets:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        df = df.where(pd.notnull(df), None)
        data["pages"][sheet_name] = df.to_dict(orient='records')

    return data

def generate_dashboard_from_excel(excel_file='procurement_data_v3.xlsx', output_file='procurement_dashboard_v3_from_excel.html'):
    """從 Excel 讀取數據並生成儀表板"""

    # 讀取 Excel 數據
    data = load_data_from_excel(excel_file)
    pages = data["pages"]

    # 取得各頁面數據
    summary_kpi = pages.get('決策總覽', [])
    alerts = pages.get('決策總覽_警示', [])
    spend_category = pages.get('支出類別', [])
    spend_region = pages.get('支出區域', [])
    spend_supplier = pages.get('支出供應商', [])
    suppliers = pages.get('供應商', [])
    cost_type = pages.get('節約類型', [])
    cost_target = pages.get('節約目標', [])
    risk_items = pages.get('風險項目', [])
    esg_scores = pages.get('ESG評分', [])
    esg_carbon = pages.get('碳排放', [])
    quality_spc = pages.get('SPC數據', [])
    quality_defect = pages.get('品質異常', [])
    contracts = pages.get('合約', [])
    maturity = pages.get('成熟度維度', [])

    # 準備圖表數據（避免 f-string 轉義問題）
    spend_category_labels = json.dumps([s['類別'] for s in spend_category])
    spend_category_values = json.dumps([s['金額'] for s in spend_category])
    spend_region_x = json.dumps([r['區域'] for r in spend_region])
    spend_region_y = json.dumps([r['金額'] for r in spend_region])
    supplier_risk = json.dumps([s['風險'] for s in suppliers])
    supplier_perf = json.dumps([s['績效'] for s in suppliers])
    supplier_names = json.dumps([s['名稱'] for s in suppliers])
    cost_type_labels = json.dumps([c['類型'] for c in cost_type])
    cost_type_values = json.dumps([c['金額'] for c in cost_type])
    quality_spc_weeks = json.dumps([q['週次'] for q in quality_spc])
    quality_spc_cpk = json.dumps([q['CPK'] for q in quality_spc])
    quality_spc_target = json.dumps([q['目標'] for q in quality_spc])
    maturity_scores = json.dumps([m['分數'] for m in maturity])
    maturity_dims = json.dumps([m['維度'] for m in maturity])

    # 構建 HTML
    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>採購分析 V3.0 - XYZ 製造</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        * {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Segoe UI', 'Microsoft JhengHei', Arial, sans-serif; background: #0f172a; color: #f8fafc; }}
        .header {{ background: linear-gradient(135deg, #1e293b, #334155); padding: 20px 30px; border-bottom: 1px solid rgba(255,255,255,0.1); }}
        .header h1 {{ font-size: 24px; color: #38bdf8; margin-bottom: 5px; }}
        .header .subtitle {{ font-size: 14px; color: #94a3b8; }}
        .nav {{ background: #1e293b; padding: 15px 0; display: flex; gap: 10px; flex-wrap: wrap; justify-content: center; }}
        .nav a {{ color: #94a3b8; text-decoration: none; padding: 10px 20px; border-radius: 8px; transition: 0.3s; font-size: 14px; }}
        .nav a:hover, .nav a.active {{ background: #3b82f6; color: white; }}
        .page {{ display: none; padding: 30px; max-width: 1600px; margin: 0 auto; }}
        .page.active {{ display: block; }}
        .card {{ background: #1e293b; border-radius: 12px; padding: 20px; margin-bottom: 20px; border: 1px solid rgba(255,255,255,0.05); }}
        .card h2 {{ color: #38bdf8; font-size: 18px; margin-bottom: 15px; border-bottom: 2px solid #3b82f6; padding-bottom: 10px; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin: 20px 0; }}
        .kpi-box {{ background: linear-gradient(135deg, #334155, #1e293b); padding: 20px; border-radius: 10px; text-align: center; }}
        .kpi-box.primary {{ background: linear-gradient(135deg, #3b82f6, #1d4ed8); }}
        .kpi-box.success {{ background: linear-gradient(135deg, #10b981, #059669); }}
        .kpi-value {{ font-size: 28px; font-weight: bold; }}
        .kpi-label {{ font-size: 12px; opacity: 0.9; margin-top: 5px; }}
        .kpi-delta {{ font-size: 12px; margin-top: 5px; }}
        .kpi-delta.up {{ color: #10b981; }}
        .kpi-delta.down {{ color: #ef4444; }}
        .chart-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 20px; margin: 20px 0; }}
        .chart-container {{ background: #1e293b; border-radius: 10px; padding: 15px; min-height: 300px; }}
        table {{ width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px; }}
        th {{ background: #334155; color: #94a3b8; padding: 12px; text-align: left; border-bottom: 1px solid #475569; }}
        td {{ padding: 10px 12px; border-bottom: 1px solid #334155; }}
        tr:hover {{ background: #334155; }}
        .badge {{ padding: 3px 10px; border-radius: 20px; font-size: 11px; font-weight: 600; }}
        .badge-high {{ background: rgba(239, 68, 68, 0.2); color: #ef4444; }}
        .badge-medium {{ background: rgba(245, 158, 11, 0.2); color: #f59e0b; }}
        .badge-low {{ background: rgba(16, 185, 129, 0.2); color: #10b981; }}
        .alert {{ padding: 12px 15px; border-radius: 8px; margin-bottom: 10px; display: flex; gap: 10px; align-items: center; }}
        .alert-critical {{ background: rgba(239, 68, 68, 0.15); border-left: 3px solid #ef4444; }}
        .alert-warning {{ background: rgba(245, 158, 11, 0.15); border-left: 3px solid #f59e0b; }}
        .alert-info {{ background: rgba(59, 130, 246, 0.15); border-left: 3px solid #3b82f6; }}
        .import-section {{ background: #1e293b; padding: 20px; border-radius: 10px; margin: 20px 0; text-align: center; }}
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 採購智能分析平台 V3.0</h1>
        <div class="subtitle">XYZ 製造股份有限公司 | 電子製造服務 (EMS) | 2025-Q4 | 數據來源：Excel</div>
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
    </nav>

    <!-- Page 1: 決策總覽 -->
    <div id="summary" class="page active">
        <div class="card">
            <h2>📈 關鍵績效指標 (KPI)</h2>
            <div class="kpi-grid">
                <div class="kpi-box primary">
                    <div class="kpi-value">${summary_kpi[0]['數值']:,}</div>
                    <div class="kpi-label">{summary_kpi[0]['KPI']} ({summary_kpi[0]['單位']})</div>
                    <div class="kpi-delta {'up' if summary_kpi[0]['趨勢'] == 'up' else 'down'}">{summary_kpi[0]['變化']:+.1f}% vs 去年同期</div>
                </div>
                <div class="kpi-box success">
                    <div class="kpi-value">${summary_kpi[1]['數值']:,}</div>
                    <div class="kpi-label">{summary_kpi[1]['KPI']}</div>
                    <div class="kpi-delta {'up' if summary_kpi[1]['趨勢'] == 'up' else 'down'}">{summary_kpi[1]['變化']:+.1f}%達成</div>
                </div>
                <div class="kpi-box">
                    <div class="kpi-value">{summary_kpi[3]['數值']}%</div>
                    <div class="kpi-label">{summary_kpi[3]['KPI']}</div>
                    <div class="kpi-delta {'up' if summary_kpi[3]['趨勢'] == 'up' else 'down'}">{summary_kpi[3]['變化']:+.1f}%</div>
                </div>
                <div class="kpi-box">
                    <div class="kpi-value">{summary_kpi[4]['數值']}%</div>
                    <div class="kpi-label">{summary_kpi[4]['KPI']}</div>
                    <div class="kpi-delta {'up' if summary_kpi[4]['趨勢'] == 'up' else 'down'}">{summary_kpi[4]['變化']:+.1f}%</div>
                </div>
            </div>
        </div>
        <div class="card">
            <h2>⚠️ 警示與提醒</h2>
'''

    # 添加警示訊息
    for alert in alerts:
        alert_class = 'alert-critical' if alert['類型'] == 'critical' else ('alert-warning' if alert['類型'] == 'warning' else 'alert-info')
        badge_class = 'badge-high' if alert['類型'] == 'critical' else ('badge-medium' if alert['類型'] == 'warning' else 'badge-low')
        html += f'            <div class="alert {alert_class}"><span>{alert["訊息"]}</span><span class="badge {badge_class}">{alert["動作"]}</span></div>\n'

    html += f'''        </div>
        <div class="chart-grid">
            <div class="chart-container"><div id="chart-spend-pie"></div></div>
            <div class="chart-container"><div id="chart-region-bar"></div></div>
        </div>
    </div>

    <!-- Page 2: 支出分析 -->
    <div id="spend" class="page">
        <div class="card">
            <h2>💰 支出分析</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-category"></div></div>
                <div class="chart-container"><div id="chart-region"></div></div>
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
            <table>
                <tr><th>名稱</th><th>類別</th><th>金額</th><th>績效</th><th>風險</th><th>評級</th></tr>
'''

    for s in suppliers:
        html += f'                <tr><td>{s["名稱"]}</td><td>{s["類別"]}</td><td>${s["金額"]:,}</td><td>{s["績效"]}</td><td>{s["風險"]}</td><td>{s["評級"]}</td></tr>\n'

    html += '''            </table>
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
'''

    for r in risk_items:
        severity_class = 'badge-high' if r['嚴重性'] == 'HIGH' else ('badge-medium' if r['嚴重性'] == 'MEDIUM' else 'badge-low')
        html += f'                <tr><td>{r["ID"]}</td><td>{r["類別"]}</td><td>{r["描述"]}</td><td><span class="badge {severity_class}">{r["嚴重性"]}</span></td><td>${r["影響金額"]:,}</td><td>{r["狀態"]}</td><td>{r["負責人"]}</td></tr>\n'

    html += '''            </table>
        </div>
    </div>

    <!-- Page 6: ESG -->
    <div id="esg" class="page">
        <div class="card">
            <h2>🌿 ESG 永續評核</h2>
            <div class="kpi-grid">
'''

    for e in esg_scores:
        color_class = 'success' if e['維度'] == '環境(E)' else ('warning' if e['維度'] == '社會(S)' else 'primary')
        html += f'                <div class="kpi-box {color_class}"><div class="kpi-value">{e["分數"]}</div><div class="kpi-label">{e["維度"]} 評分</div><div class="kpi-delta">目標: {e["目標"]}</div></div>\n'

    html += '''            </div>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-carbon"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 7: 品質 SPC -->
    <div id="quality" class="page">
        <div class="card">
            <h2>📏 品質 SPC 管制</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-cpk"></div></div>
                <div class="chart-container"><div id="chart-defect"></div></div>
            </div>
        </div>
    </div>

    <!-- Page 8: 合約管理 -->
    <div id="contract" class="page">
        <div class="card">
            <h2>📄 合約生命週期</h2>
            <table>
                <tr><th>合約編號</th><th>供應商</th><th>類型</th><th>金額</th><th>到期日</th><th>狀態</th></tr>
'''

    for c in contracts:
        status_class = 'badge-low' if c['狀態'] == '有效' else 'badge-high'
        html += f'                <tr><td>{c["合約編號"]}</td><td>{c["供應商"]}</td><td>{c["類型"]}</td><td>${c["金額"]:,}</td><td>{c["到期日"]}</td><td><span class="badge {status_class}">{c["狀態"]}</span></td></tr>\n'

    html += '''            </table>
        </div>
    </div>

    <!-- Page 9: 採購成熟度 -->
    <div id="maturity" class="page">
        <div class="card">
            <h2>🎓 採購成熟度評估</h2>
            <div class="chart-grid">
                <div class="chart-container"><div id="chart-maturity-radar"></div></div>
                <div class="chart-container"><div id="chart-maturity-bar"></div></div>
            </div>
        </div>
    </div>

    <script>
    function showPage(pageId) {{
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.nav a').forEach(a => a.classList.remove('active'));
        document.getElementById(pageId).classList.add('active');
        document.querySelector('a[href="#' + pageId + '"]').classList.add('active');
        renderCharts();
    }}

    const spendCategoryLabels = ''' + spend_category_labels + ''';
    const spendCategoryValues = ''' + spend_category_values + ''';
    const spendRegionX = ''' + spend_region_x + ''';
    const spendRegionY = ''' + spend_region_y + ''';
    const supplierRisk = ''' + supplier_risk + ''';
    const supplierPerf = ''' + supplier_perf + ''';
    const supplierNames = ''' + supplier_names + ''';
    const costTypeLabels = ''' + cost_type_labels + ''';
    const costTypeValues = ''' + cost_type_values + ''';
    const qualitySpcWeeks = ''' + quality_spc_weeks + ''';
    const qualitySpcCpk = ''' + quality_spc_cpk + ''';
    const qualitySpcTarget = ''' + quality_spc_target + ''';
    const maturityScores = ''' + maturity_scores + ''';
    const maturityDims = ''' + maturity_dims + ''';


    function renderCharts() {{
        // 支出類別圓餅圖
        if (document.getElementById('chart-spend-pie')) {{
            Plotly.newPlot('chart-spend-pie', [{{
                labels: spendCategoryLabels,
                values: spendCategoryValues,
                type: 'pie',
                marker: {{colors: ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6']}}
            }}], {{title: '支出類別分布', template: 'plotly_dark'}});
        }}

        // 區域長條圖
        if (document.getElementById('chart-region-bar')) {{
            Plotly.newPlot('chart-region-bar', [{{
                x: spendRegionX,
                y: spendRegionY,
                type: 'bar',
                marker: {{color: '#3b82f6'}}
            }}], {{title: '區域支出', template: 'plotly_dark'}});
        }}

        // 供應商散佈圖
        if (document.getElementById('chart-supplier-scatter')) {{
            Plotly.newPlot('chart-supplier-scatter', [{{
                x: supplierRisk,
                y: supplierPerf,
                text: supplierNames,
                mode: 'markers+text',
                marker: {{size: 15, color: '#3b82f6'}}
            }}], {{title: '供應商績效 vs 風險', xaxis: {{title: '風險分數'}}, yaxis: {{title: '績效分數'}}, template: 'plotly_dark'}});
        }}

        // 成本節約類型
        if (document.getElementById('chart-cost-type')) {{
            Plotly.newPlot('chart-cost-type', [{{
                labels: costTypeLabels,
                values: costTypeValues,
                type: 'pie',
                marker: {{colors: ['#10b981', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899']}}
            }}], {{title: '節約類型分布', template: 'plotly_dark'}});
        }}

        // CPK 趨勢圖
        if (document.getElementById('chart-cpk')) {{
            Plotly.newPlot('chart-cpk', [{{
                x: qualitySpcWeeks,
                y: qualitySpcCpk,
                type: 'scatter',
                mode: 'lines+markers',
                name: 'CPK',
                line: {{color: '#10b981'}}
            }}, {{
                x: qualitySpcWeeks,
                y: qualitySpcTarget,
                type: 'scatter',
                mode: 'lines',
                name: '目標',
                line: {{color: '#ef4444', dash: 'dash'}}
            }}], {{title: 'CPK 趨勢', template: 'plotly_dark'}});
        }}

        // 成熟度雷達圖
        if (document.getElementById('chart-maturity-radar')) {{
            Plotly.newPlot('chart-maturity-radar', [{{
                type: 'scatterpolar',
                r: maturityScores,
                theta: maturityDims,
                fill: 'toself',
                marker: {{color: '#3b82f6'}}
            }}], {{title: '成熟度雷達', polar: {{radialaxis: {{visible: true, range: [0, 5]}}}}, template: 'plotly_dark'}});
        }}
    }}

    renderCharts();
    </script>
</body>
</html>'''

    # 寫入 HTML 檔案
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"儀表板生成成功！")
    print(f"輸出檔案: {output_file}")
    print(f"工作表數量: {len(pages)}")
    for page_name in pages:
        print(f"  - {page_name}: {len(pages[page_name])} 筆資料")

    return data

if __name__ == "__main__":
    try:
        data = generate_dashboard_from_excel()
        print("\nExcel 數據讀取並生成儀表板成功！")
    except FileNotFoundError:
        print("Excel 文件不存在，請先執行 create_excel_template.py 創建模板")
