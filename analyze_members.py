"""
analyze_members.py
------------------
会員購買データを4つのSheetに集計して出力する。
時間軸は PICKUP_TIME 列を基準とする。

  Sheet1「概要」        : 全体 vs 会員の金額サマリ（月別内訳含む）
  Sheet2「会員別集計」  : 会員IDごとの購買統計
  Sheet3「商品別集計」  : 会員購買における商品別数量・金額・割合
  Sheet4「月別推移」    : PICKUP_TIME 基準の月別会員購買推移

用法：
    python analyze_members.py                              # 自動でカレントの .xlsx を読込
    python analyze_members.py 1月.xlsx 2月.xlsx 3月.xlsx
    python analyze_members.py --dir ./data --output 会員分析Q1.xlsx
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter


# ── 設定 ────────────────────────────────────────────────────
DEFAULT_OUTPUT   = "会員分析.xlsx"
DEFAULT_SHEET_IN = 0

# 列名（ソースファイルの表頭と一致させること）
COL_MEMBER      = "会員ID"
COL_POS         = "POS番号"
COL_RECEIPT     = "レシート番号"
COL_PRODUCT     = "商品名"
COL_PROD_CODE   = "商品コード"
COL_QTY         = "数量"
COL_AMOUNT      = "税込金額"
COL_PICKUP_TIME = "PICKUP_TIME"   # 時間軸の基準列
COL_TXN         = "__txn_key__"   # POS番号＋レシート番号 の複合キー

HEADER_BG = "1F4E79"
HEADER_FG = "FFFFFF"
# ─────────────────────────────────────────────────────────────


def collect_files(files, directory):
    if files:
        paths = [Path(f) for f in files]
    elif directory:
        paths = sorted(Path(directory).glob("*.xlsx"))
    else:
        paths = sorted(Path(".").glob("*.xlsx"))

    paths = [p for p in paths if p.resolve() != Path(DEFAULT_OUTPUT).resolve()]
    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        sys.exit(f"[エラー] ファイルが見つかりません: {', '.join(missing)}")
    if not paths:
        sys.exit("[エラー] .xlsx ファイルが見つかりません。")
    return paths


def read_one_file(p, sheet_index):
    """1ファイルを読み込む。openpyxl で失敗した場合は xlrd（.xls 形式）でリトライ。"""
    try:
        return pd.read_excel(p, sheet_name=sheet_index, keep_default_na=False, engine="openpyxl")
    except KeyError as e:
        if "[Content_Types].xml" in str(e):
            # 旧形式 .xls の可能性 → xlrd で再試行
            print(f"    ⚠  openpyxl で読込失敗。旧形式(.xls)として再試行: {p.name}")
            try:
                return pd.read_excel(p, sheet_name=sheet_index, keep_default_na=False, engine="xlrd")
            except Exception as e2:
                sys.exit(f"[エラー] '{p.name}' を読み込めません: {e2}\n"
                         f"  ヒント: Excel で一度開いて「名前を付けて保存」→「.xlsx 形式」で保存し直してください。")
        raise


def load_all(paths, sheet_index):
    """全ファイルを読み込んで1つの DataFrame に結合する。
    PICKUP_TIME を datetime に変換し、年月列（__ym__）を付与する。
    """
    dfs = []
    for p in paths:
        print(f"  読込中: {p.name}")
        df = read_one_file(p, sheet_index)
        df["__source__"] = p.stem
        dfs.append(df)
    df_all = pd.concat(dfs, ignore_index=True)

    # POS番号＋レシート番号 の複合トランザクションキーを生成
    df_all[COL_TXN] = df_all[COL_POS].astype(str) + "_" + df_all[COL_RECEIPT].astype(str)

    # PICKUP_TIME → datetime → 年月ラベル（例: "2025-01"）
    df_all[COL_PICKUP_TIME] = pd.to_datetime(df_all[COL_PICKUP_TIME], errors="coerce")
    df_all["__ym__"] = df_all[COL_PICKUP_TIME].dt.to_period("M").astype(str)

    print(f"  合計: {len(df_all):,} 行")
    return df_all


def build_summary(df_all):
    """Sheet1: 全体 vs 会員 の金額サマリ（縦型）。"""
    total_amount   = df_all[COL_AMOUNT].sum()
    df_member      = df_all[df_all[COL_MEMBER].astype(str).str.strip() != ""]
    member_count   = df_member[COL_MEMBER].nunique()
    member_txn     = df_member[COL_TXN].nunique()
    member_freq    = member_txn / member_count if member_count else 0
    member_amount  = df_member[COL_AMOUNT].sum()
    ratio          = member_amount / total_amount if total_amount else 0

    rows = [
        ("全体購入金額合計（税込）",       total_amount),
        ("会員総人数",                     member_count),
        ("会員トランザクション数",          member_txn),
        ("会員購入回数（平均/人）",         round(member_freq, 2)),
        ("会員購入金額合計（税込）",        member_amount),
        ("会員金額占比",                   ratio),
    ]
    return pd.DataFrame(rows, columns=["項目", "値"]), df_member


def build_member_stats(df_member):
    """Sheet2: 会員IDごとの購買統計。"""
    grp = df_member.groupby(COL_MEMBER, sort=False)
    stats = pd.DataFrame({
        "購入回数":           grp[COL_TXN].nunique(),
        "購入点数":           grp[COL_QTY].sum(),
        "購入金額合計（税込）": grp[COL_AMOUNT].sum(),
    }).reset_index()
    stats["平均購入金額/回"] = (
        stats["購入金額合計（税込）"] / stats["購入回数"]
    ).round(1)
    return stats.sort_values("購入金額合計（税込）", ascending=False).reset_index(drop=True)


def build_monthly_stats(df_all, df_member):
    """Sheet4: PICKUP_TIME 基準の月別推移（全体 vs 会員）。"""
    total_amount = df_all[COL_AMOUNT].sum()

    # 全体月別
    grp_all = df_all.groupby("__ym__")
    monthly_all = pd.DataFrame({
        "全体購入金額":   grp_all[COL_AMOUNT].sum(),
        "全体トランザクション数": grp_all[COL_TXN].nunique(),
    }).reset_index().rename(columns={"__ym__": "年月"})

    # 会員月別
    grp_mem = df_member.groupby("__ym__")
    monthly_mem = pd.DataFrame({
        "会員トランザクション数": grp_mem[COL_TXN].nunique(),
        "会員購入回数（平均/人）": (grp_mem[COL_TXN].nunique() / grp_mem[COL_MEMBER].nunique()).round(2),
        "会員購入点数":           grp_mem[COL_QTY].sum(),
        "会員購入金額":           grp_mem[COL_AMOUNT].sum(),
        "会員人数（月）":         grp_mem[COL_MEMBER].nunique(),
    }).reset_index().rename(columns={"__ym__": "年月"})

    stats = monthly_all.merge(monthly_mem, on="年月", how="left").fillna(0)
    stats["会員金額占比"] = stats["会員購入金額"] / stats["全体購入金額"].replace(0, float("nan"))
    return stats.sort_values("年月").reset_index(drop=True)


def build_product_stats(df_member):
    """Sheet3: 商品別 数量・金額・割合（会員購買のみ）。"""
    total_member_amount = df_member[COL_AMOUNT].sum()

    grp = df_member.groupby([COL_PROD_CODE, COL_PRODUCT], sort=False)
    stats = pd.DataFrame({
        "購入数量":  grp[COL_QTY].sum(),
        "購入金額":  grp[COL_AMOUNT].sum(),
    }).reset_index()
    stats["金額割合"] = stats["購入金額"] / total_member_amount if total_member_amount else 0
    return stats.sort_values("購入金額", ascending=False).reset_index(drop=True)


# ── スタイル適用 ──────────────────────────────────────────────

def style_sheet(ws, pct_cols=None):
    """表頭スタイル・列幅・冻結・数値フォーマットを一括適用。"""
    pct_cols = pct_cols or []

    h_font  = Font(bold=True, color=HEADER_FG, size=10)
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.font  = h_font
        cell.fill  = h_fill
        cell.alignment = h_align

    # 割合列に % フォーマット
    for col_letter in pct_cols:
        for row in range(2, ws.max_row + 1):
            ws[f"{col_letter}{row}"].number_format = "0.00%"

    # 列幅
    for col_idx, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1):
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"


def write_excel(output_path, df_summary, df_members, df_products, df_monthly):
    print(f"\n  書込中: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        # ── Sheet1: 概要 ──────────────────────────────────────
        df_summary.to_excel(writer, sheet_name="概要", index=False)
        ws1 = writer.sheets["概要"]
        style_sheet(ws1)
        ws1["B7"].number_format = "0.00%"  # 会員金額占比（6行目データ = row 7）

        # ── Sheet2: 会員別集計 ────────────────────────────────
        df_members.to_excel(writer, sheet_name="会員別集計", index=False)
        ws2 = writer.sheets["会員別集計"]
        style_sheet(ws2)
        print(f"    [会員別集計] {len(df_members):,} 会員")

        # ── Sheet3: 商品別集計 ────────────────────────────────
        df_products.to_excel(writer, sheet_name="商品別集計", index=False)
        ws3 = writer.sheets["商品別集計"]
        pct_col = get_column_letter(df_products.columns.get_loc("金額割合") + 1)
        style_sheet(ws3, pct_cols=[pct_col])
        print(f"    [商品別集計] {len(df_products):,} 商品")

        # ── Sheet4: 月別推移（PICKUP_TIME 基準）──────────────
        df_monthly.to_excel(writer, sheet_name="月別推移", index=False)
        ws4 = writer.sheets["月別推移"]
        pct_col4 = get_column_letter(df_monthly.columns.get_loc("会員金額占比") + 1)
        style_sheet(ws4, pct_cols=[pct_col4])
        print(f"    [月別推移] {len(df_monthly)} ヶ月分")

    print(f"  完成！→ {output_path}")


def main():
    parser = argparse.ArgumentParser(description="会員購買データ集計ツール")
    parser.add_argument("files", nargs="*", help="対象 Excel ファイル（省略時は自動スキャン）")
    parser.add_argument("--dir",      default=None,           help="スキャン対象ディレクトリ")
    parser.add_argument("--output",   default=DEFAULT_OUTPUT, help=f"出力ファイル名（デフォルト: {DEFAULT_OUTPUT}）")
    parser.add_argument("--sheet-in", type=int, default=DEFAULT_SHEET_IN, help="読込 Sheet インデックス（0起算）")
    args = parser.parse_args()

    print("=== 会員購買分析ツール ===")
    paths = collect_files(args.files, args.dir)
    print(f"対象ファイル: {len(paths)} 件")

    df_all = load_all(paths, args.sheet_in)

    print("\n  集計中...")
    df_summary, df_member = build_summary(df_all)
    df_members  = build_member_stats(df_member)
    df_products = build_product_stats(df_member)
    df_monthly  = build_monthly_stats(df_all, df_member)

    write_excel(args.output, df_summary, df_members, df_products, df_monthly)


if __name__ == "__main__":
    main()
