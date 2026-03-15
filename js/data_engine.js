// data_engine.js
// 負責資料載入、解析與全局狀態管理

const AppData = {
    summary: {
        sum: 8400000, sumTrend: 5.2,
        saving: 4.8, savingTarget: 4.0,
        otd: 92.4, otdTrend: -1.2,
        highRiskSuppliers: 3,
        trendLabels: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        trendOTD: [94.5, 95.1, 93.8, 92.0, 93.6, 92.4],
        trendQuality: [98.2, 98.5, 98.0, 97.5, 97.9, 97.1],
        alerts: [
            { type: '交期延遲', target: '聯創材料', severity: 'danger', status: '待處理' },
            { type: '價格異常', target: '精華原液 B', severity: 'warning', status: '調查中' },
            { type: '品質客訴', target: '華美包裝', severity: 'danger', status: '已發出整改' }
        ]
    },
    spend: {
        categories: ['原物料', '包材', '設備', '服務', '其他'],
        categoryValues: [4500, 1800, 1200, 600, 300],
        suppliers: ['榮達原料', '美源科技', '天澤化工', '星研美業', '嘉博原料'],
        supplierValues: [2200, 1800, 1500, 1100, 800],
        maverickLabels: ['IT硬體', '辦公耗材', '緊急零組件', '行銷外包'],
        maverickValues: [120, 85, 210, 150]
    },
    supplier: {
        quadrant: [
            { x: 97.8, y: 99.3, r: 15, name: '榮達原料' },
            { x: 93.8, y: 98.8, r: 18, name: '天澤化工' },
            { x: 88.8, y: 98.1, r: 10, name: '嘉博原料' },
            { x: 96.8, y: 99.2, r: 25, name: '美源科技' },
            { x: 91.8, y: 99.0, r: 12, name: '星研美業' },
            { x: 79.8, y: 96.9, r: 8,  name: '博瑞成分' }, // High Risk
            { x: 94.8, y: 97.6, r: 9,  name: '華美包裝' },
            { x: 90.8, y: 98.9, r: 7,  name: '創新生物' },
            { x: 82.8, y: 97.6, r: 6,  name: '聯創材料' }  // High Risk
        ],
        topBottomLabels: ['榮達', '美源', '天澤', '聯創', '博瑞'],
        topBottomScores: [95.1, 91.5, 93.7, 74.2, 55.3]
    },
    cost: {
        avoidance: 1200000,
        ppv: 150000,
        roi: 420,
        months: ['Q1', 'Q2', 'Q3', 'Q4'],
        priceCurrent: [100, 105, 108, 112],
        priceLastYear: [95, 98, 100, 102],
        priceMarket: [102, 110, 115, 120]
    },
    quality: {
        defectLabels: ['尺寸偏差', '外觀瑕疵', '成分超標', '包裝破損', '標籤錯誤'],
        defectValues: [45, 30, 15, 8, 2],
        deliveryLabels: ['提前到貨', '準時', '延遲 1-3 天', '延遲 >3天'],
        deliveryValues: [15, 65, 12, 8]
    },
    deviation: {
        catLabels: ['精華液', '洗護', '香氛', '彩妝'],
        planVals: [1000, 800, 600, 1200],
        actualVals: [1150, 750, 600, 1100], // 超採精華液, 漏採洗護, 彩妝
        trendMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
        trendVals: [98, 95, 102, 88, 92]
    },
    efficiency: {
        radarLabels: ['採購部', '研發部', '生產部', '行銷部', '行政部'],
        radarEmergency: [5, 15, 25, 40, 15],
        radarEfficiency: [90, 85, 75, 60, 80]
    }
};

// Handle file upload via SheetJS
document.getElementById('fileInput').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    document.getElementById('data-status').innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> 資料解析中...`;
    
    const reader = new FileReader();
    reader.onload = function(evt) {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, {type: 'binary'});
            
            // Assume we have a specific sheet format, or we just mock a refresh to show interaction
            console.log("Sheet names:", workbook.SheetNames);
            
            // Apply mock changes for demonstration when a file is uploaded to simulate real-time effect
            AppData.summary.sum = 9200000;
            AppData.summary.saving = 5.1;
            AppData.summary.otd = 94.2;
            AppData.summary.highRiskSuppliers = 1;
            
            // Refresh Dashboard UI components
            if (window.refreshDashboard) {
                window.refreshDashboard();
            }

            document.getElementById('data-status').innerHTML = `<i class="fa-solid fa-check text-green"></i> 數據已更新 (${file.name})`;
        } catch (error) {
            console.error(error);
            document.getElementById('data-status').innerHTML = `<i class="fa-solid fa-triangle-exclamation text-red"></i> 讀取錯誤`;
        }
    };
    reader.readAsBinaryString(file);
});
