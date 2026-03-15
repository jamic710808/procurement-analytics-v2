// dashboard.js
// 負責頁面切換、圖表初始化與全網互動渲染

function applyThemeDefaults() {
    const isLight = document.body.getAttribute('data-theme') === 'light';
    Chart.defaults.color = isLight ? '#475569' : '#94a3b8';
    Chart.defaults.font.family = "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";
    Chart.defaults.scale.grid.color = isLight ? 'rgba(0, 0, 0, 0.05)' : 'rgba(255, 255, 255, 0.05)';
    Chart.defaults.plugins.tooltip.backgroundColor = isLight ? 'rgba(255, 255, 255, 0.9)' : 'rgba(15, 23, 42, 0.9)';
    Chart.defaults.plugins.tooltip.titleColor = isLight ? '#0f172a' : '#fff';
    Chart.defaults.plugins.tooltip.bodyColor = isLight ? '#475569' : '#fff';
    Chart.defaults.plugins.tooltip.padding = 10;
    Chart.defaults.plugins.tooltip.borderColor = isLight ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.1)';
    Chart.defaults.plugins.tooltip.borderWidth = 1;
}
applyThemeDefaults();

let charts = {};

// Navigation Logic
document.querySelectorAll('.nav-links a').forEach(link => {
    link.addEventListener('click', (e) => {
        e.preventDefault();
        // Update active class on nav
        document.querySelectorAll('.nav-links a').forEach(n => n.classList.remove('active'));
        e.currentTarget.classList.add('active');
        
        // Hide all pages
        document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
        
        // Show target page
        const targetId = e.currentTarget.getAttribute('data-target');
        document.getElementById(targetId).classList.add('active');
        
        // Update Header Title
        document.getElementById('current-page-title').innerText = e.currentTarget.innerText;
    });
});

// Initialization
function initDashboard() {
    applyThemeDefaults();
    renderSummaryPage();
    renderSpendPage();
    renderSupplierPage();
    renderCostPage();
    renderQualityPage();
    renderDeviationPage();
    renderEfficiencyPage();
    renderSimulatorPage();
    window.runSimulation(); // Init simulator
}

// 1. Executive Summary
function renderSummaryPage() {
    // KPI Updates
    document.getElementById('kpi-sum').innerText = '$' + (AppData.summary.sum / 1000000).toFixed(1) + 'M';
    document.getElementById('kpi-saving').innerText = AppData.summary.saving + '%';
    document.getElementById('kpi-otd').innerText = AppData.summary.otd + '%';
    document.getElementById('kpi-risk').innerText = AppData.summary.highRiskSuppliers;

    // Trend Chart
    const ctxTrend = document.getElementById('chart-summary-trend').getContext('2d');
    if(charts.summaryTrend) charts.summaryTrend.destroy();
    charts.summaryTrend = new Chart(ctxTrend, {
        type: 'line',
        data: {
            labels: AppData.summary.trendLabels,
            datasets: [
                {
                    label: '準時交付率 (OTD)',
                    data: AppData.summary.trendOTD,
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    tension: 0.4,
                    fill: true
                },
                {
                    label: '綜合合格率',
                    data: AppData.summary.trendQuality,
                    borderColor: '#10b981',
                    borderDash: [5, 5],
                    tension: 0.4
                }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    // Alert Table
    const alertBody = document.querySelector('#table-summary-alerts tbody');
    alertBody.innerHTML = '';
    AppData.summary.alerts.forEach(alert => {
        const badgeClass = alert.severity === 'danger' ? 'status-danger' : 'status-warning';
        alertBody.innerHTML += `
            <tr>
                <td>${alert.type}</td>
                <td style="color:var(--text-primary)">${alert.target}</td>
                <td><span class="status-badge ${badgeClass}">${alert.severity.toUpperCase()}</span></td>
                <td>${alert.status}</td>
            </tr>
        `;
    });
}

// 2. Spend Analysis
function renderSpendPage() {
    const ctxCat = document.getElementById('chart-spend-category').getContext('2d');
    if(charts.spendCat) charts.spendCat.destroy();
    charts.spendCat = new Chart(ctxCat, {
        type: 'doughnut',
        data: {
            labels: AppData.spend.categories,
            datasets: [{
                data: AppData.spend.categoryValues,
                backgroundColor: ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#64748b'],
                borderWidth:0
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, cutout: '70%' }
    });

    const ctxSup = document.getElementById('chart-spend-supplier').getContext('2d');
    if(charts.spendSup) charts.spendSup.destroy();
    charts.spendSup = new Chart(ctxSup, {
        type: 'bar',
        data: {
            labels: AppData.spend.suppliers,
            datasets: [{
                label: '支出金額 (K)',
                data: AppData.spend.supplierValues,
                backgroundColor: '#3b82f6',
                borderRadius: 4
            }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const ctxMav = document.getElementById('chart-spend-maverick').getContext('2d');
    if(charts.spendMav) charts.spendMav.destroy();
    charts.spendMav = new Chart(ctxMav, {
        type: 'bar',
        data: {
            labels: AppData.spend.maverickLabels,
            datasets: [{
                label: '非正規採購支出金額 (K)',
                data: AppData.spend.maverickValues,
                backgroundColor: '#ef4444',
                borderRadius: 4
            }]
        },
        options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false }
    });
}

// 3. Supplier Risk
function renderSupplierPage() {
    const ctxQuad = document.getElementById('chart-supplier-quadrant').getContext('2d');
    if(charts.supQuad) charts.supQuad.destroy();
    
    // Process bubble colors
    const bubbleData = AppData.supplier.quadrant.map(d => ({
        x: d.x, y: d.y, r: d.r, name: d.name,
        // If x < 90 or y < 98 -> Warning/Danger
        bgColor: (d.x < 90 && d.y < 98) ? 'rgba(239, 68, 68, 0.6)' : 
                 (d.x < 95 || d.y < 98) ? 'rgba(245, 158, 11, 0.6)' : 
                                          'rgba(16, 185, 129, 0.6)'
    }));

    charts.supQuad = new Chart(ctxQuad, {
        type: 'bubble',
        data: {
            datasets: [{
                label: '供應商',
                data: bubbleData,
                backgroundColor: bubbleData.map(d => d.bgColor),
                borderColor: bubbleData.map(d => d.bgColor.replace('0.6', '1.0')),
                borderWidth: 1
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                x: { title: { display: true, text: '平均準時率 (%)'}, min: 75, max: 100 },
                y: { title: { display: true, text: '平均合格率 (%)'}, min: 95, max: 100 }
            },
            plugins: {
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const d = context.raw;
                            return `${d.name}: 準時 ${d.x}%, 合格 ${d.y}% (佔比規模: ${d.r})`;
                        }
                    }
                }
            }
        }
    });

    const ctxRank = document.getElementById('chart-supplier-ranking').getContext('2d');
    if(charts.supRank) charts.supRank.destroy();
    charts.supRank = new Chart(ctxRank, {
        type: 'bar',
        data: {
            labels: AppData.supplier.topBottomLabels,
            datasets: [{
                label: '綜合評分',
                data: AppData.supplier.topBottomScores,
                backgroundColor: AppData.supplier.topBottomScores.map(score => score > 80 ? '#10b981' : '#ef4444'),
                borderRadius: 4
            }]
        },
        options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false }
    });
}

// 4. Cost Saving
function renderCostPage() {
    const ctxPrice = document.getElementById('chart-cost-price-variance').getContext('2d');
    if(charts.costPrice) charts.costPrice.destroy();
    charts.costPrice = new Chart(ctxPrice, {
        type: 'line',
        data: {
            labels: AppData.cost.months,
            datasets: [
                { label: '本年採購價', data: AppData.cost.priceCurrent, borderColor: '#3b82f6', tension: 0.3 },
                { label: '去年採購價', data: AppData.cost.priceLastYear, borderColor: '#64748b', borderDash: [5,5], tension: 0.3 },
                { label: '市場均價', data: AppData.cost.priceMarket, borderColor: '#ef4444', tension: 0.3 }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });
}

// 5. Quality
function renderQualityPage() {
    const ctxDefect = document.getElementById('chart-quality-defects').getContext('2d');
    if(charts.qualDefect) charts.qualDefect.destroy();
    charts.qualDefect = new Chart(ctxDefect, {
        type: 'bar',
        data: {
            labels: AppData.quality.defectLabels,
            datasets: [{
                label: '缺陷發生次數',
                data: AppData.quality.defectValues,
                backgroundColor: '#8b5cf6',
                borderRadius: 4
            }]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const ctxDel = document.getElementById('chart-quality-delivery').getContext('2d');
    if(charts.qualDel) charts.qualDel.destroy();
    charts.qualDel = new Chart(ctxDel, {
        type: 'doughnut',
        data: {
            labels: AppData.quality.deliveryLabels,
            datasets: [{
                data: AppData.quality.deliveryValues,
                backgroundColor: ['#10b981', '#3b82f6', '#f59e0b', '#ef4444'],
                borderWidth:0
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, cutout: '70%' }
    });
}

// 6. Deviation
function renderDeviationPage() {
    const ctxCat = document.getElementById('chart-deviation-category').getContext('2d');
    if(charts.devCat) charts.devCat.destroy();
    charts.devCat = new Chart(ctxCat, {
        type: 'bar',
        data: {
            labels: AppData.deviation.catLabels,
            datasets: [
                { label: '計畫量', data: AppData.deviation.planVals, backgroundColor: '#64748b' },
                { label: '實際採購量', data: AppData.deviation.actualVals, backgroundColor: '#3b82f6' }
            ]
        },
        options: { responsive: true, maintainAspectRatio: false }
    });

    const ctxTrend = document.getElementById('chart-deviation-trend').getContext('2d');
    if(charts.devTrend) charts.devTrend.destroy();
    charts.devTrend = new Chart(ctxTrend, {
        type: 'line',
        data: {
            labels: AppData.deviation.trendMonths,
            datasets: [{
                label: '月度計畫完成率 (%)',
                data: AppData.deviation.trendVals,
                borderColor: '#10b981',
                tension: 0.4
            }]
        },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 50, max: 120 } } }
    });
}

// 7. Efficiency
function renderEfficiencyPage() {
    const ctxRadar = document.getElementById('chart-efficiency-radar').getContext('2d');
    if(charts.effRadar) charts.effRadar.destroy();
    charts.effRadar = new Chart(ctxRadar, {
        type: 'radar',
        data: {
            labels: AppData.efficiency.radarLabels,
            datasets: [
                {
                    label: '緊急採購佔比 (%)',
                    data: AppData.efficiency.radarEmergency,
                    backgroundColor: 'rgba(239, 68, 68, 0.2)',
                    borderColor: '#ef4444'
                },
                {
                    label: '流程自動化/效率分 (0-100)',
                    data: AppData.efficiency.radarEfficiency,
                    backgroundColor: 'rgba(59, 130, 246, 0.2)',
                    borderColor: '#3b82f6'
                }
            ]
        },
        options: { 
            responsive: true, 
            maintainAspectRatio: false, 
            scales: { 
                r: { 
                    grid: { color: document.body.getAttribute('data-theme') === 'light' ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.1)' }, 
                    angleLines: { color: document.body.getAttribute('data-theme') === 'light' ? 'rgba(0,0,0,0.1)' : 'rgba(255,255,255,0.1)' } 
                } 
            } 
        }
    });
}

// 8. Simulator
window.runSimulation = function() {
    const matInput = parseFloat(document.getElementById('sim-mat').value);
    const delayInput = parseFloat(document.getElementById('sim-delay').value);
    const demandInput = parseFloat(document.getElementById('sim-demand').value);

    // Mock Business Logic Simulation
    let baseCost = 5000000;
    let newCost = baseCost * (1 + (matInput / 100)) * (1 + (demandInput / 100));
    // Delay inherently increases cost slightly due to expedited shipping needs
    newCost += delayInput * 15000; 

    // Inventory Gap (%) = Demand increases but delivery delayed
    let gap = (demandInput * 0.8) + (delayInput * 2) - 5; 
    if(gap < 0) gap = 0;

    const ctxRes = document.getElementById('chart-simulator-result').getContext('2d');
    if(charts.simRes) charts.simRes.destroy();
    charts.simRes = new Chart(ctxRes, {
        type: 'bar',
        data: {
            labels: ['當前基準', '模擬預測結果'],
            datasets: [
                {
                    label: '預估總成本 ($)',
                    data: [baseCost, newCost],
                    backgroundColor: ['#64748b', newCost > baseCost ? '#ef4444' : '#10b981'],
                    yAxisID: 'y'
                },
                {
                    label: '潛在庫存缺口風險 (%)',
                    data: [5, gap],
                    type: 'line',
                    borderColor: '#f59e0b',
                    yAxisID: 'y1'
                }
            ]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            scales: {
                y: { type: 'linear', position: 'left', title: {display:true, text:'成本'} },
                y1: { type: 'linear', position: 'right', grid: {drawOnChartArea: false}, title: {display:true, text:'風險 %'}, min: 0, max: 100 }
            }
        }
    });
};

// Theme Toggle Logic
document.addEventListener('DOMContentLoaded', () => {
    const themeBtn = document.getElementById('theme-toggle');
    const themeIcon = document.getElementById('theme-icon');
    const themeText = document.getElementById('theme-text');

    if(themeBtn) {
        themeBtn.addEventListener('click', () => {
            const isLight = document.body.getAttribute('data-theme') === 'light';
            if(isLight) {
                document.body.removeAttribute('data-theme');
                themeIcon.className = 'fa-solid fa-sun';
                themeText.innerText = '白天模式';
            } else {
                document.body.setAttribute('data-theme', 'light');
                themeIcon.className = 'fa-solid fa-moon';
                themeText.innerText = '夜間模式';
            }
            initDashboard(); // Re-render all charts and UI updates
        });
    }
});

// Expose refresh to data engine
window.refreshDashboard = function() {
    initDashboard(); // Ensure full reinit if needed
};

// Bootstrap
document.addEventListener('DOMContentLoaded', initDashboard);
