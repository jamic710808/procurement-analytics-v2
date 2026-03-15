import csv
import random
from datetime import timedelta, date

suppliers = ['供應商A', '供應商B', '台積電', '聯發科', '大立光', '鴻海', '廣達']
categories = ['電子零組件', '機構件', '包材', '間接物料']
depts = ['RD', 'IT', 'Manufacturing', 'Marketing']

data = []
start_date = date(2024, 1, 1)

headers = ['日期', '供應商', '品類', '部門', '採購金額', '節約金額', '品質良率', '交期準確率', '風險評分', 'ESG評分']

with open('c:/Users/jamic/採購分析/sample_data_v2.csv', 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(headers)
    for i in range(50):
        dt = start_date + timedelta(days=random.randint(0, 729))
        spend = random.randint(500, 5000) * 1000
        writer.writerow([
            dt.strftime('%Y-%m-%d'),
            random.choice(suppliers),
            random.choice(categories),
            random.choice(depts),
            spend,
            int(spend * random.uniform(0.01, 0.1)),
            round(random.uniform(80, 100), 1),
            round(random.uniform(70, 100), 1),
            round(random.uniform(5, 50), 1),
            round(random.uniform(50, 95), 1)
        ])
