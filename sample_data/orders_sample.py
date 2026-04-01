"""実データ仕様に合わせたサンプルExcelを生成するスクリプト"""
import pandas as pd
from pathlib import Path

rows = [
    # order_id, customer_id, Product_id, quantity, sales, order_date, Separate
    # カート1: 他社IPのみ
    ("ORD-001", "C001", "P001", 1, 3000, "2024-01-05", "他社IP"),
    ("ORD-001", "C001", "P-BAG", 1, 500,  "2024-01-05", "shoppingbag"),
    # カート2: 自社IPのみ
    ("ORD-002", "C002", "P010", 2, 6000, "2024-01-06", "自社IP"),
    # カート3: 他社IP + 自社IP（クロス）
    ("ORD-003", "C003", "P001", 1, 3000, "2024-01-07", "他社IP"),
    ("ORD-003", "C003", "P010", 1, 3000, "2024-01-07", "自社IP"),
    ("ORD-003", "C003", "P-BAG", 1, 500,  "2024-01-07", "shoppingbag"),
    # カート4: 他社IPのみ（非会員：customer_idなし）
    ("ORD-004", None,   "P002", 1, 2500, "2024-01-08", "他社IP"),
    # カート5: 自社IPのみ
    ("ORD-005", "C005", "P011", 1, 3500, "2024-01-09", "自社IP"),
    # カート6: 自社IPのみ（非会員）
    ("ORD-006", None,   "P012", 3, 9000, "2024-01-10", "自社IP"),
    # カート7: 他社IP + 自社IP（クロス）
    ("ORD-007", "C007", "P002", 1, 2500, "2024-02-01", "他社IP"),
    ("ORD-007", "C007", "P011", 1, 3500, "2024-02-01", "自社IP"),
    # カート8: 他社IPのみ
    ("ORD-008", "C008", "P001", 2, 6000, "2024-02-03", "他社IP"),
    ("ORD-008", "C008", "P-BAG", 1, 500,  "2024-02-03", "shoppingbag"),
    # カート9: 自社IPのみ
    ("ORD-009", "C009", "P010", 1, 3000, "2024-02-05", "自社IP"),
    # カート10: 他社IPのみ（非会員）
    ("ORD-010", None,   "P003", 1, 2000, "2024-02-06", "他社IP"),
]

df = pd.DataFrame(rows, columns=[
    "order_id", "customer_id", "Product_id",
    "quantity", "sales", "order_date", "Separate"
])

Path("sample_data").mkdir(exist_ok=True)
df.to_excel("sample_data/orders_sample.xlsx", index=False)
print(f"生成完了: sample_data/orders_sample.xlsx ({len(df)}行)")
