"""
CrossUse（コラボ×自社IP）セグメント分析
=========================================
注文ID単位で、以下の3セグメントに分類して集計します：
  - collab_only   : コラボ商品のみ購入
  - cross         : 同一注文でコラボ+自社IP両方購入
  - own_ip_only   : 自社IP商品のみ購入

【必要なファイル】
  orders.csv   : 取引データ（order_id, customer_id, product_id, ...）
  products.csv : 商品マスタ（product_id, product_type='collab'/'own_ip', ...）

【使い方】
  python analyze_crossuse.py
  # → results/crossuse_analysis.xlsx に結果を出力
"""

import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── ファイルパス設定（実データに合わせて変更してください） ──────────────
ORDERS_FILE   = "sample_data/orders_sample.csv"   # 取引データ
PRODUCTS_FILE = "sample_data/products_sample.csv" # 商品マスタ
OUTPUT_DIR    = "results"
OUTPUT_FILE   = f"{OUTPUT_DIR}/crossuse_analysis.xlsx"

# 列名の設定（実データの列名に合わせて変更）
ORDER_ID_COL   = "order_id"
PRODUCT_ID_COL = "product_id"
TYPE_COL       = "product_type"   # 商品マスタの区分列
COLLAB_VALUE   = "collab"         # コラボ商品の値
OWN_IP_VALUE   = "own_ip"         # 自社IP商品の値


# ────────────────────────────────────────────────
# 1. データ読み込み
# ────────────────────────────────────────────────
def load_data(orders_path: str, products_path: str):
    def read_file(path):
        p = Path(path)
        if p.suffix in (".xlsx", ".xls"):
            return pd.read_excel(p)
        return pd.read_csv(p)

    orders   = read_file(orders_path)
    products = read_file(products_path)

    print(f"[読込] 取引データ: {len(orders):,}行  商品マスタ: {len(products):,}行")
    return orders, products


# ────────────────────────────────────────────────
# 2. 注文ID単位でセグメント分類
# ────────────────────────────────────────────────
def classify_orders(orders: pd.DataFrame, products: pd.DataFrame) -> pd.DataFrame:
    # 商品マスタと結合して各行に product_type を付与
    df = orders.merge(
        products[[PRODUCT_ID_COL, TYPE_COL]],
        on=PRODUCT_ID_COL,
        how="left"
    )

    # マスタに存在しない商品のチェック
    missing = df[TYPE_COL].isna().sum()
    if missing > 0:
        print(f"[警告] 商品マスタに存在しない商品が {missing} 行あります（除外して集計します）")
        df = df.dropna(subset=[TYPE_COL])

    # 注文ID単位で「コラボを含むか」「自社IPを含むか」をフラグ化
    grouped = df.groupby(ORDER_ID_COL)[TYPE_COL]
    order_flags = pd.DataFrame({
        "has_collab": grouped.apply(lambda t: (t == COLLAB_VALUE).any()),
        "has_own_ip": grouped.apply(lambda t: (t == OWN_IP_VALUE).any()),
    }).reset_index()

    # セグメント分類
    def segment(row):
        if row["has_collab"] and row["has_own_ip"]:
            return "cross"          # コラボ + 自社IP 同時購入
        elif row["has_collab"]:
            return "collab_only"    # コラボのみ
        else:
            return "own_ip_only"    # 自社IPのみ

    order_flags["segment"] = order_flags.apply(segment, axis=1)
    return order_flags


# ────────────────────────────────────────────────
# 3. 集計
# ────────────────────────────────────────────────
def summarize(order_flags: pd.DataFrame) -> pd.DataFrame:
    total = len(order_flags)

    summary = (
        order_flags.groupby("segment")
        .size()
        .reset_index(name="注文数")
    )

    # 表示順を固定
    segment_order = {
        "collab_only":  "コラボのみ",
        "cross":        "コラボ + 自社IP（同時購入）",
        "own_ip_only":  "自社IPのみ",
    }
    summary["セグメント名"] = summary["segment"].map(segment_order)
    summary["構成比(%)"]   = (summary["注文数"] / total * 100).round(1)
    summary = summary[["セグメント名", "注文数", "構成比(%)"]]

    # 合計行を追加
    total_row = pd.DataFrame([{
        "セグメント名": "合計",
        "注文数":       total,
        "構成比(%)":    100.0,
    }])
    summary = pd.concat([summary, total_row], ignore_index=True)

    print("\n=== CrossUseセグメント集計結果 ===")
    print(summary.to_string(index=False))
    print()
    return summary


# ────────────────────────────────────────────────
# 4. Excel出力（見やすいフォーマット）
# ────────────────────────────────────────────────
BLUE_HEADER = "1F4E79"
COLORS = {
    "collab_only":  "F4B942",   # オレンジ：コラボのみ
    "cross":        "70AD47",   # グリーン：クロス購買
    "own_ip_only":  "4472C4",   # ブルー：自社IPのみ
    "total":        "595959",   # グレー：合計
}
SEG_COLOR_MAP = {
    "コラボのみ":                COLORS["collab_only"],
    "コラボ + 自社IP（同時購入）": COLORS["cross"],
    "自社IPのみ":                COLORS["own_ip_only"],
    "合計":                      COLORS["total"],
}


def _thin_border():
    s = Side(border_style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def export_excel(summary: pd.DataFrame, order_flags: pd.DataFrame, output_path: str):
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()

    # ── シート1: サマリ ──────────────────────────────────
    ws = wb.active
    ws.title = "セグメント集計"
    ws.sheet_view.showGridLines = False

    # タイトル
    ws.merge_cells("A1:C1")
    c = ws["A1"]
    c.value = "CrossUse分析 ― コラボ×自社IP セグメント集計"
    c.font = Font(bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ヘッダー
    headers = ["セグメント名", "注文数", "構成比(%)"]
    col_widths = [30, 12, 14]
    for i, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(2, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # データ行
    for row_idx, row in summary.iterrows():
        r = row_idx + 3
        seg_name = row["セグメント名"]
        color    = SEG_COLOR_MAP.get(seg_name, "CCCCCC")

        # セグメント名
        c = ws.cell(r, 1, seg_name)
        c.font = Font(bold=(seg_name == "合計"), size=10)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)

        # 注文数
        c = ws.cell(r, 2, row["注文数"])
        c.number_format = '#,##0'
        c.alignment = Alignment(horizontal="right")
        c.fill = PatternFill("solid", fgColor="F2F2F2" if seg_name != "合計" else "DDDDDD")
        if seg_name == "合計":
            c.font = Font(bold=True)

        # 構成比
        c = ws.cell(r, 3, row["構成比(%)"] / 100)
        c.number_format = '0.0%'
        c.alignment = Alignment(horizontal="right")
        c.fill = PatternFill("solid", fgColor="F2F2F2" if seg_name != "合計" else "DDDDDD")
        if seg_name == "合計":
            c.font = Font(bold=True)

        ws.row_dimensions[r].height = 22

    # ボーダー
    for r in range(2, 2 + len(summary) + 1):
        for col in range(1, 4):
            ws.cell(r, col).border = _thin_border()

    # ── シート2: 注文IDリスト ────────────────────────────
    ws2 = wb.create_sheet("注文セグメント一覧")
    ws2.sheet_view.showGridLines = False

    ws2.merge_cells("A1:C1")
    c = ws2["A1"]
    c.value = "注文ID別 セグメント分類一覧"
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 24

    seg_label_map = {
        "collab_only": "コラボのみ",
        "cross":       "コラボ + 自社IP（同時購入）",
        "own_ip_only": "自社IPのみ",
    }
    detail_headers = [ORDER_ID_COL, "コラボ含む", "自社IP含む", "セグメント"]
    detail_widths  = [20, 12, 12, 30]
    for i, (h, w) in enumerate(zip(detail_headers, detail_widths), 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
        c = ws2.cell(2, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 18

    for row_idx, row in order_flags.reset_index(drop=True).iterrows():
        r = row_idx + 3
        seg_key   = row["segment"]
        seg_label = seg_label_map[seg_key]
        color     = COLORS[seg_key]

        ws2.cell(r, 1, row[ORDER_ID_COL]).alignment = Alignment(horizontal="center")
        ws2.cell(r, 2, "○" if row["has_collab"] else "―").alignment = Alignment(horizontal="center")
        ws2.cell(r, 3, "○" if row["has_own_ip"] else "―").alignment = Alignment(horizontal="center")

        c = ws2.cell(r, 4, seg_label)
        c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.alignment = Alignment(horizontal="left", indent=1)
        ws2.row_dimensions[r].height = 16

        for col in range(1, 5):
            ws2.cell(r, col).border = _thin_border()

    wb.save(output_path)
    print(f"[出力] {output_path}")


# ────────────────────────────────────────────────
# メイン実行
# ────────────────────────────────────────────────
if __name__ == "__main__":
    orders, products = load_data(ORDERS_FILE, PRODUCTS_FILE)
    order_flags      = classify_orders(orders, products)
    summary          = summarize(order_flags)
    export_excel(summary, order_flags, OUTPUT_FILE)
