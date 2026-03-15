import pandas as pd
import numpy as np

def generate_sample_data():
    """
    產生採購分析儀表板的模擬數據 (Excel 格式)
    """
    print("正在產生範例數據...")

    # 1. 供應商績效表 (Supplier Performance)
    suppliers = ['榮達原料', '天澤化工', '嘉博原料', '美源科技', '星研美業', '博瑞成分', '華美包裝', '創新生物', '聯創材料']
    supplier_data = pd.DataFrame({
        '供應商名稱': suppliers,
        '採購金額_K': np.random.randint(500, 3000, size=9),
        '平均準時率_%': np.random.uniform(75.0, 99.5, size=9).round(1),
        '平均合格率_%': np.random.uniform(95.0, 99.9, size=9).round(1),
        '風險評級': np.random.choice(['低', '中', '高'], size=9, p=[0.6, 0.3, 0.1])
    })

    # 2. 支出類別表 (Spend Categories)
    categories = ['原物料', '包材', '設備', '服務', '其他']
    spend_data = pd.DataFrame({
        '支出類別': categories,
        '本季支出_K': [4500, 1800, 1200, 600, 300],
        '合約覆蓋率_%': [95, 88, 70, 40, 20]
    })

    # 3. 價格差異追蹤表 (Price Variance)
    materials = ['精華原液A', '穩定劑B', '環保包裝C']
    months = ['Q1', 'Q2', 'Q3', 'Q4']
    
    price_records = []
    for m in materials:
        last_year = np.random.randint(80, 120)
        for q in months:
            current = int(last_year * np.random.uniform(0.95, 1.15))
            market = int(current * np.random.uniform(0.9, 1.1))
            price_records.append({
                '原物料': m,
                '季度': q,
                '去年採購價': last_year,
                '本年採購價': current,
                '市場均價': market
            })
    price_data = pd.DataFrame(price_records)

    # 4. 品質異常與交期異常紀錄 (Quality & Delivery Issues)
    issue_data = pd.DataFrame({
        '異常類別': ['尺寸偏差', '外觀瑕疵', '成分超標', '包裝破損', '標籤錯誤'],
        '發生次數': [45, 30, 15, 8, 2],
        '影響金額_K': [120, 80, 200, 15, 5]
    })

    # 輸出至 Excel (多 Worksheets)
    output_file = 'Procurement_Sample_Data.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        supplier_data.to_excel(writer, sheet_name='供應商績效', index=False)
        spend_data.to_excel(writer, sheet_name='類別支出', index=False)
        price_data.to_excel(writer, sheet_name='價格差異追蹤', index=False)
        issue_data.to_excel(writer, sheet_name='品質客訴紀錄', index=False)

    print(f"✅ 成功產出範例資料表：{output_file}")

if __name__ == "__main__":
    generate_sample_data()
