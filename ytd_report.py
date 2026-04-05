"""
ytd_report.py
-------------
YTD（Year-To-Date）累積レポートを5シートで出力する。
全月のデータを合算し、年間累計の KPI を集計する。

  Sheet1「概要」  : YTD 全体 KPI（縦型）
  Sheet2「月別」  : 月ごとの KPI（横並び）
  Sheet3「店舗別」: 店舗別 YTD KPI
  Sheet4「IP別」  : IP別 YTD KPI（IP マスタ必須）
  Sheet5「商品別」: 商品別 YTD KPI

KPI:
  全体購入金額合計（税込）| TXN数 | 点数 | 連帯率 | 客単価
  免税購入金額合計（税込）| 免税比率 | 免税TXN数 | 免税数量 | 免税連帯率 | 免税客単価
  会員金額（税込）| 会員人数 | 会員客単価 | 会員連帯率 | 会員購入頻度

用法:
    python ytd_report.py 1月.xlsx 2月.xlsx 3月.xlsx
    python ytd_report.py --dir ./data --output YTD_Q1.xlsx
    python ytd_report.py --ip-file IP対応表.xlsx
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ── 列名定義 ─────────────────────────────────────────────────
COL_POS         = "POS番号"
COL_RECEIPT     = "レシート番号"
COL_MEMBER      = "会員ID"
COL_PRODUCT     = "商品名"
COL_PROD_CODE   = "商品コード"
COL_QTY         = "数量"
COL_AMOUNT      = "税込金額"
COL_TAX         = "税金"
COL_STORE       = "店舗"
COL_PICKUP_TIME = "PICKUP_TIME"
COL_TXN         = "__txn_key__"
COL_IP          = "IP名称"

# ── 設定 ─────────────────────────────────────────────────────
DEFAULT_OUTPUT = "YTDレポート.xlsx"
SHEET_IN       = 0
OKU            = 100_000_000

TAX_EXEMPT_METHOD      = "tax_zero"   # "tax_zero" or "type_col"
COL_TYPE               = "TYPE"
TAX_EXEMPT_TYPE_VALUES = {"免税", "TAX_FREE", "TAXFREE", "DUTY_FREE"}
RETURN_TYPE_VALUES     = {"Return", "RETURN", "返品", "返却", "REFUND"}
TOTAL_KEYWORDS         = r"合計|小計|総計|total|subtotal|grand"

HEADER_BG  = "1F4E79"
HEADER_FG  = "FFFFFF"
# ─────────────────────────────────────────────────────────────

# KPI の定義順（概要の縦並び・各シートの列順に使用）
KPI_ORDER = [
    "全体購入金額合計（税込）",
    "TXN数",
    "点数",
    "連帯率",
    "客単価",
    "免税購入金額合計（税込）",
    "免税比率",
    "免税TXN数",
    "免税数量",
    "免税連帯率",
    "免税客単価",
    "会員金額（税込）",
    "会員人数",
    "会員客単価",
    "会員連帯率",
    "会員購入頻度",
]

PCT_KPI  = {"免税比率"}                             # % フォーマット
OKU_KPI  = {"全体購入金額合計（税込）",              # 億円換算（概要のみ）
             "免税購入金額合計（税込）",
             "会員金額（税込）"}


# ══════════════════════════════════════════════════════════════
#  データ読込
# ══════════════════════════════════════════════════════════════

def collect_files(files, directory, output):
    if files:
        paths = [Path(f) for f in files]
    elif directory:
        paths = sorted(Path(directory).glob("*.xlsx"))
    else:
        paths = sorted(Path(".").glob("*.xlsx"))

    out_abs = Path(output).resolve()
    paths = [p for p in paths
             if p.resolve() != out_abs and "ip" not in p.stem.lower()]

    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        sys.exit(f"[エラー] 見つかりません: {', '.join(missing)}")
    if not paths:
        sys.exit("[エラー] .xlsx ファイルが見つかりません。")
    return paths


def read_one_file(p):
    try:
        return pd.read_excel(p, sheet_name=SHEET_IN, keep_default_na=False, engine="openpyxl")
    except KeyError as e:
        if "[Content_Types].xml" in str(e):
            print(f"    ⚠ 旧形式として再試行: {p.name}")
            try:
                return pd.read_excel(p, sheet_name=SHEET_IN, keep_default_na=False, engine="xlrd")
            except Exception as e2:
                sys.exit(f"[エラー] '{p.name}' 読込失敗: {e2}")
        raise


def load_all(paths):
    dfs = []
    for p in paths:
        print(f"  読込中: {p.name}")
        df = read_one_file(p)
        df["__source__"] = p.stem
        dfs.append(df)
    df_all = pd.concat(dfs, ignore_index=True)

    # 集計行除外（POS番号の空白は有効取引のため除外しない）
    key_empty = (
        df_all[COL_RECEIPT].astype(str).str.strip().isin(["", "nan"]) |
        df_all[COL_PROD_CODE].astype(str).str.strip().isin(["", "nan"])
    )
    kw_match = (
        df_all[COL_PRODUCT].astype(str).str.contains(TOTAL_KEYWORDS, case=False, na=False) |
        df_all[COL_PROD_CODE].astype(str).str.contains(TOTAL_KEYWORDS, case=False, na=False)
    )
    removed = (key_empty | kw_match).sum()
    if removed:
        print(f"  ⚠ 集計行 {removed:,} 行除外")
    df_all = df_all[~(key_empty | kw_match)].reset_index(drop=True)

    # TXN 複合キー
    df_all[COL_TXN] = df_all[COL_POS].astype(str) + "_" + df_all[COL_RECEIPT].astype(str)

    # PICKUP_TIME → 年月
    df_all[COL_PICKUP_TIME] = pd.to_datetime(df_all[COL_PICKUP_TIME], errors="coerce")
    df_all["__ym__"] = df_all[COL_PICKUP_TIME].dt.to_period("M").astype(str)
    invalid = df_all[COL_PICKUP_TIME].isna()
    if invalid.sum():
        print(f"  ⚠ PICKUP_TIME 無効 {invalid.sum():,} 行 → '日付不明'")
    df_all.loc[invalid, "__ym__"] = "日付不明"

    # 返品行補正
    if COL_TYPE in df_all.columns:
        is_return = df_all[COL_TYPE].astype(str).str.strip().isin(RETURN_TYPE_VALUES)
        if is_return.sum():
            for col in [COL_AMOUNT, COL_QTY]:
                df_all.loc[is_return & (df_all[col] > 0), col] *= -1
            print(f"  返品行: {is_return.sum():,} 行  返品金額: {df_all.loc[is_return, COL_AMOUNT].sum():,.0f} 円")

    # 免税フラグ
    if TAX_EXEMPT_METHOD == "type_col" and COL_TYPE in df_all.columns:
        df_all["__exempt__"] = df_all[COL_TYPE].astype(str).str.upper().isin(
            {v.upper() for v in TAX_EXEMPT_TYPE_VALUES}
        )
    else:
        df_all["__exempt__"] = pd.to_numeric(df_all[COL_TAX], errors="coerce").fillna(0) == 0

    print(f"  有効行: {len(df_all):,} 行")
    return df_all


def load_ip_master(ip_file):
    if ip_file:
        path = Path(ip_file)
    else:
        candidates = [p for p in sorted(Path(".").glob("*.xlsx"))
                      if "ip" in p.stem.lower() or "IP" in p.stem]
        if not candidates:
            return None
        path = candidates[0]
    print(f"  IP マスタ: {path.name}")
    df = pd.read_excel(path, header=0, keep_default_na=False,
                       engine="openpyxl", usecols=[0, 1])
    df.columns = [COL_PROD_CODE, COL_IP]
    df[COL_PROD_CODE] = df[COL_PROD_CODE].astype(str).str.strip()
    df[COL_IP]        = df[COL_IP].astype(str).str.strip()
    return df.drop_duplicates(subset=COL_PROD_CODE)


# ══════════════════════════════════════════════════════════════
#  KPI 計算
# ══════════════════════════════════════════════════════════════

def calc_kpi(df) -> dict:
    """任意の DataFrame スライスに対して全 KPI を計算して dict で返す。"""
    # 全体
    txn    = df[COL_TXN].nunique()
    qty    = df[COL_QTY].sum()
    amount = df[COL_AMOUNT].sum()

    # 免税
    df_ex  = df[df["__exempt__"]]
    ex_txn = df_ex[COL_TXN].nunique()
    ex_qty = df_ex[COL_QTY].sum()
    ex_amt = df_ex[COL_AMOUNT].sum()

    # 会員
    df_m    = df[df[COL_MEMBER].astype(str).str.strip() != ""]
    m_count = df_m[COL_MEMBER].nunique()
    m_txn   = df_m[COL_TXN].nunique()
    m_qty   = df_m[COL_QTY].sum()
    m_amt   = df_m[COL_AMOUNT].sum()

    return {
        "全体購入金額合計（税込）": amount,
        "TXN数":                 txn,
        "点数":                  qty,
        "連帯率":                round(qty    / txn,     2) if txn     else 0,
        "客単価":                round(amount / txn,     1) if txn     else 0,
        "免税購入金額合計（税込）": ex_amt,
        "免税比率":              ex_amt / amount          if amount  else 0,
        "免税TXN数":             ex_txn,
        "免税数量":              ex_qty,
        "免税連帯率":            round(ex_qty / ex_txn,  2) if ex_txn  else 0,
        "免税客単価":            round(ex_amt / ex_txn,  1) if ex_txn  else 0,
        "会員金額（税込）":       m_amt,
        "会員人数":              m_count,
        "会員客単価":            round(m_amt  / m_count, 1) if m_count else 0,
        "会員連帯率":            round(m_qty  / m_count, 2) if m_count else 0,
        "会員購入頻度":          round(m_txn  / m_count, 2) if m_count else 0,
    }


# ══════════════════════════════════════════════════════════════
#  Sheet ビルド
# ══════════════════════════════════════════════════════════════

def build_overview(df_all):
    """Sheet1: YTD 全体 KPI（縦型）。金額 KPI は億円換算。"""
    kpi = calc_kpi(df_all)
    rows = []
    for key in KPI_ORDER:
        val = kpi[key]
        if key in OKU_KPI:
            rows.append({"指標": f"{key}[億円]", "YTD": round(val / OKU, 4)})
        else:
            rows.append({"指標": key, "YTD": val})
    return pd.DataFrame(rows)


def build_monthly(df_all):
    """Sheet2: 月別 KPI（年月 | KPI列…）。"""
    months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    rows = []
    for ym in months:
        kpi = calc_kpi(df_all[df_all["__ym__"] == ym])
        rows.append({"年月": ym, **{k: kpi[k] for k in KPI_ORDER}})
    return pd.DataFrame(rows)


def _build_by_dim(df_all, group_keys):
    """任意のグループキーで集計した KPI テーブルを返す。"""
    groups = df_all.groupby(group_keys, sort=False)
    records = []
    for keys, grp in groups:
        kpi = calc_kpi(grp)
        if isinstance(keys, tuple):
            row = dict(zip(group_keys, keys))
        else:
            row = {group_keys[0]: keys}
        row.update({k: kpi[k] for k in KPI_ORDER})
        records.append(row)

    df = pd.DataFrame(records)
    return df.sort_values("全体購入金額合計（税込）", ascending=False).reset_index(drop=True)


def build_by_store(df_all):
    """Sheet3: 店舗別 YTD KPI。"""
    return _build_by_dim(df_all, [COL_STORE])


def build_by_ip(df_all, df_ip_master):
    """Sheet4: IP別 YTD KPI。"""
    df = df_all.copy()
    df[COL_PROD_CODE] = df[COL_PROD_CODE].astype(str).str.strip()
    df = df.merge(df_ip_master[[COL_PROD_CODE, COL_IP]], on=COL_PROD_CODE, how="left")
    df[COL_IP] = df[COL_IP].fillna("IP未設定")
    return _build_by_dim(df, [COL_IP])


def build_by_product(df_all):
    """Sheet5: 商品別 YTD KPI。"""
    return _build_by_dim(df_all, [COL_PROD_CODE, COL_PRODUCT])


# ══════════════════════════════════════════════════════════════
#  スタイル & 出力
# ══════════════════════════════════════════════════════════════

def style_sheet(ws, df):
    h_font  = Font(bold=True, color=HEADER_FG, size=10)
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

    # % フォーマット列を特定
    for i, col in enumerate(df.columns, 1):
        letter = get_column_letter(i)
        is_pct = any(k in col for k in ("比率", "構成比", "占比"))
        if is_pct:
            for row in range(2, ws.max_row + 1):
                ws[f"{letter}{row}"].number_format = "0.00%"

    # 列幅
    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1
    ):
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"


def write_excel(output_path, sheets: dict):
    print(f"\n  書込中: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            style_sheet(writer.sheets[name[:31]], df)
            print(f"    [{name}] {len(df):,} 行")
    print(f"\n  完成！→ {output_path}")


# ══════════════════════════════════════════════════════════════
#  エントリポイント
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="YTD レポート生成ツール")
    parser.add_argument("files",     nargs="*", help="対象 xlsx（省略時は自動スキャン）")
    parser.add_argument("--dir",     default=None,           help="スキャンディレクトリ")
    parser.add_argument("--output",  default=DEFAULT_OUTPUT, help="出力ファイル名")
    parser.add_argument("--ip-file", default=None,           help="SKU→IP マスタ")
    args = parser.parse_args()

    print("=== YTD レポート生成ツール ===")
    paths = collect_files(args.files, args.dir, args.output)
    print(f"対象: {len(paths)} ファイル")

    df_all       = load_all(paths)
    df_ip_master = load_ip_master(args.ip_file)

    valid_months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    print(f"  集計期間: {valid_months[0]} ～ {valid_months[-1]}  ({len(valid_months)} ヶ月)")

    print("  集計中...")
    sheets = {
        "概要":  build_overview(df_all),
        "月別":  build_monthly(df_all),
        "店舗別": build_by_store(df_all),
        "商品別": build_by_product(df_all),
    }
    if df_ip_master is not None:
        sheets = dict(list(sheets.items())[:3]) | \
                 {"IP別": build_by_ip(df_all, df_ip_master)} | \
                 {"商品別": sheets["商品別"]}
    else:
        print("  ⚠ IP マスタなし → IP別シートをスキップ")

    write_excel(args.output, sheets)


if __name__ == "__main__":
    main()
