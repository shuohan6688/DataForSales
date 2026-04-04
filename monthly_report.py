"""
monthly_report.py
-----------------
月次レポートを 5 シートで出力する。

  Sheet1「概要」  : 当月 KPI + 前月 + 先月比（縦型）
  Sheet2「月別」  : 全月の推移 + 先月比
  Sheet3「店舗別」: 店舗別集計 + 構成比 + 先月比
  Sheet4「IP別」  : IP別集計 + 構成比 + 先月比（IP マスタ必須）
  Sheet5「商品別」: 商品別集計 + 構成比 + 先月比

当月 = 入力データの最新月、前月 = その 1 か月前（データがあれば）

用法:
    python monthly_report.py 1月.xlsx 2月.xlsx 3月.xlsx
    python monthly_report.py --dir ./data --output 月次レポート.xlsx
    python monthly_report.py --ip-file IP対応表.xlsx
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
DEFAULT_OUTPUT = "月次レポート.xlsx"
SHEET_IN       = 0
OKU            = 100_000_000   # 億円

# 免税判定: "tax_zero" → 税金列が 0 の行、"type_col" → TYPE列の値で判定
TAX_EXEMPT_METHOD      = "tax_zero"
COL_TYPE               = "TYPE"
TAX_EXEMPT_TYPE_VALUES = {"免税", "TAX_FREE", "TAXFREE", "DUTY_FREE"}
RETURN_TYPE_VALUES     = {"Return", "RETURN", "返品", "返却", "REFUND"}  # 返品タイプ

TOTAL_KEYWORDS = r"合計|小計|総計|total|subtotal|grand"

HEADER_BG    = "1F4E79"
HEADER_FG    = "FFFFFF"
MOM_POS_BG   = "E2EFDA"   # 先月比プラス（薄緑）
MOM_NEG_BG   = "FCE4D6"   # 先月比マイナス（薄赤）
SECTION_BG   = "BDD7EE"   # 概要の区切り行
# ─────────────────────────────────────────────────────────────


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
             if p.resolve() != out_abs
             and "ip" not in p.stem.lower()]   # IP マスタは除外

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
    """全ファイル結合 → 集計行除外 → TXNキー → 年月 → 免税フラグ。"""
    dfs = []
    for p in paths:
        print(f"  読込中: {p.name}")
        df = read_one_file(p)
        df["__source__"] = p.stem
        dfs.append(df)
    df_all = pd.concat(dfs, ignore_index=True)

    # 集計行除外
    key_empty = (
        df_all[COL_POS].astype(str).str.strip().isin(["", "nan"]) |
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

    # TXN 複合キー（POS番号＋レシート番号）
    df_all[COL_TXN] = df_all[COL_POS].astype(str) + "_" + df_all[COL_RECEIPT].astype(str)

    # PICKUP_TIME → 年月
    df_all[COL_PICKUP_TIME] = pd.to_datetime(df_all[COL_PICKUP_TIME], errors="coerce")
    df_all["__ym__"] = df_all[COL_PICKUP_TIME].dt.to_period("M").astype(str)
    invalid = df_all[COL_PICKUP_TIME].isna()
    if invalid.sum():
        print(f"  ⚠ PICKUP_TIME 無効 {invalid.sum():,} 行 → '日付不明'")
    df_all.loc[invalid, "__ym__"] = "日付不明"

    # 返品行の金額・数量を負に補正（TYPE == Return かつ正の値の場合のみ）
    if COL_TYPE in df_all.columns:
        is_return = df_all[COL_TYPE].astype(str).str.strip().isin(RETURN_TYPE_VALUES)
        n_return  = is_return.sum()
        if n_return > 0:
            for col in [COL_AMOUNT, COL_QTY]:
                df_all.loc[is_return & (df_all[col] > 0), col] *= -1
            return_amt = df_all.loc[is_return, COL_AMOUNT].sum()
            print(f"  返品行: {n_return:,} 行  返品金額合計: {return_amt:,.0f} 円")

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
#  指標計算ヘルパー
# ══════════════════════════════════════════════════════════════

METRIC_ORDER = [
    "購入金額合計（税込）",
    "トランザクション数",
    "購入点数",
    "連帯率",
    "客単価",
    "免税購入金額合計（税込）",
    "免税比率",
]

def metrics_for(df) -> dict:
    txn    = df[COL_TXN].nunique()
    qty    = df[COL_QTY].sum()
    amount = df[COL_AMOUNT].sum()
    exempt = df.loc[df["__exempt__"], COL_AMOUNT].sum()
    return {
        "購入金額合計（税込）":     amount,
        "トランザクション数":      txn,
        "購入点数":               qty,
        "連帯率":                round(qty  / txn,    2) if txn    else 0,
        "客単価":                round(amount / txn,  1) if txn    else 0,
        "免税購入金額合計（税込）": exempt,
        "免税比率":               exempt / amount       if amount  else 0,
    }


def mom(curr, prev):
    """先月比 (curr - prev) / |prev|。前月 0 または None は None を返す。"""
    try:
        if prev is not None and prev != 0:
            return (curr - prev) / abs(prev)
    except Exception:
        pass
    return None


# ══════════════════════════════════════════════════════════════
#  Sheet ビルド
# ══════════════════════════════════════════════════════════════

def build_overview(df_curr, df_prev, curr_ym, prev_ym):
    """Sheet1: 縦型 KPI（指標 / 当月 / 前月 / 先月比）。金額は億円換算。"""
    AMT_KEYS = {"購入金額合計（税込）", "免税購入金額合計（税込）"}
    PCT_KEYS  = {"免税比率"}

    curr = metrics_for(df_curr)
    prev = metrics_for(df_prev) if df_prev is not None else {}

    rows = []
    for key in METRIC_ORDER:
        c = curr[key]
        p = prev.get(key)
        m = mom(c, p)

        if key in AMT_KEYS:
            label  = f"{key}[億円]"
            c_disp = round(c / OKU, 4)
            p_disp = round(p / OKU, 4) if p is not None else None
        else:
            label  = key
            c_disp = c
            p_disp = p

        rows.append({
            "指標":              label,
            curr_ym:            c_disp,
            prev_ym or "前月":  p_disp,
            "先月比":           m,
        })

    return pd.DataFrame(rows)


def build_monthly(df_all):
    """Sheet2: 月別推移（年月 | 各指標 | 先月比 … 交互配置）。"""
    months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")

    rows = []
    prev_m = {}
    for ym in months:
        df_ym  = df_all[df_all["__ym__"] == ym]
        curr_m = metrics_for(df_ym)
        row    = {"年月": ym}
        for key in METRIC_ORDER:
            row[key]          = curr_m[key]
            row[f"{key}_先月比"] = mom(curr_m[key], prev_m.get(key))
        rows.append(row)
        prev_m = curr_m

    # 列順: 年月, (指標, 先月比) × 7
    ordered = ["年月"]
    for key in METRIC_ORDER:
        ordered += [key, f"{key}_先月比"]
    df = pd.DataFrame(rows)
    return df[[c for c in ordered if c in df.columns]]


def _agg_by(df, group_keys):
    """グループキーで基本集計し DataFrame を返す。"""
    grp        = df.groupby(group_keys, sort=False)
    exempt_grp = df[df["__exempt__"]].groupby(group_keys, sort=False)

    base = pd.DataFrame({
        "購入金額合計（税込）":     grp[COL_AMOUNT].sum(),
        "トランザクション数":      grp[COL_TXN].nunique(),
        "購入点数":               grp[COL_QTY].sum(),
        "免税購入金額合計（税込）": exempt_grp[COL_AMOUNT].sum(),
    }).reset_index().fillna(0)

    base["連帯率"] = (
        base["購入点数"] / base["トランザクション数"].replace(0, float("nan"))
    ).round(2)
    base["客単価"] = (
        base["購入金額合計（税込）"] / base["トランザクション数"].replace(0, float("nan"))
    ).round(1)
    base["免税比率"] = (
        base["免税購入金額合計（税込）"] / base["購入金額合計（税込）"].replace(0, float("nan"))
    )
    return base


def _add_mom_cols(curr_df, df_prev, group_keys, mom_targets):
    """前月データをマージして先月比列を追加する。"""
    if df_prev is None:
        return curr_df
    prev_df = _agg_by(df_prev, group_keys)
    merged  = curr_df.merge(prev_df, on=group_keys, how="left", suffixes=("", "_前月"))
    for col in mom_targets:
        merged[f"{col}_先月比"] = merged.apply(
            lambda r, c=col: mom(r[c], r.get(f"{c}_前月")), axis=1
        )
    return merged.drop(columns=[c for c in merged.columns if c.endswith("_前月")])


def build_by_store(df_curr, df_prev):
    """Sheet3: 店舗別。"""
    t_amt = df_curr[COL_AMOUNT].sum()
    t_txn = df_curr[COL_TXN].nunique()

    df = _agg_by(df_curr, [COL_STORE])
    df["金額構成比"] = df["購入金額合計（税込）"] / t_amt if t_amt else 0
    df["TXN構成比"]  = df["トランザクション数"]  / t_txn if t_txn else 0

    df = _add_mom_cols(df, df_prev, [COL_STORE],
                       ["購入金額合計（税込）", "トランザクション数",
                        "購入点数", "免税購入金額合計（税込）"])

    col_order = [
        COL_STORE,
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "TXN構成比",  "トランザクション数_先月比",
        "購入点数", "連帯率", "客単価",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.sort_values("購入金額合計（税込）", ascending=False).reset_index(drop=True)


def build_by_ip(df_curr, df_prev, df_ip_master):
    """Sheet4: IP別。"""
    def add_ip(df):
        d = df.copy()
        d[COL_PROD_CODE] = d[COL_PROD_CODE].astype(str).str.strip()
        d = d.merge(df_ip_master[[COL_PROD_CODE, COL_IP]], on=COL_PROD_CODE, how="left")
        d[COL_IP] = d[COL_IP].fillna("IP未設定")
        return d

    dc = add_ip(df_curr)
    dp = add_ip(df_prev) if df_prev is not None else None

    t_amt = dc[COL_AMOUNT].sum()
    t_qty = dc[COL_QTY].sum()

    df = _agg_by(dc, [COL_IP])
    df["金額構成比"] = df["購入金額合計（税込）"] / t_amt if t_amt else 0
    df["数量構成比"] = df["購入点数"]            / t_qty if t_qty else 0

    df = _add_mom_cols(df, dp, [COL_IP],
                       ["購入金額合計（税込）", "トランザクション数",
                        "購入点数", "免税購入金額合計（税込）"])

    col_order = [
        COL_IP,
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "トランザクション数_先月比",
        "購入点数", "数量構成比", "購入点数_先月比",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.sort_values("購入金額合計（税込）", ascending=False).reset_index(drop=True)


def build_by_product(df_curr, df_prev):
    """Sheet5: 商品別。"""
    t_amt = df_curr[COL_AMOUNT].sum()
    t_qty = df_curr[COL_QTY].sum()

    df = _agg_by(df_curr, [COL_PROD_CODE, COL_PRODUCT])
    df["金額構成比"] = df["購入金額合計（税込）"] / t_amt if t_amt else 0
    df["数量構成比"] = df["購入点数"]            / t_qty if t_qty else 0

    df = _add_mom_cols(df, df_prev, [COL_PROD_CODE, COL_PRODUCT],
                       ["購入金額合計（税込）", "トランザクション数",
                        "購入点数", "免税購入金額合計（税込）"])

    col_order = [
        COL_PROD_CODE, COL_PRODUCT,
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "トランザクション数_先月比",
        "購入点数", "数量構成比", "購入点数_先月比",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.sort_values("購入金額合計（税込）", ascending=False).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
#  スタイル & 出力
# ══════════════════════════════════════════════════════════════

def _col_letters(df):
    """列名 → Excelの列レター の dict を返す。"""
    return {col: get_column_letter(i + 1) for i, col in enumerate(df.columns)}


def style_sheet(ws, df):
    """表頭スタイル・列幅・フォーマット・先月比色分けを適用。"""
    h_font  = Font(bold=True, color=HEADER_FG, size=10)
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

    letters = _col_letters(df)
    pct_keywords = ("比率", "構成比", "占比", "先月比")

    for col, letter in letters.items():
        is_pct   = any(k in col for k in pct_keywords)
        is_mom   = "先月比" in col

        for row in range(2, ws.max_row + 1):
            cell = ws[f"{letter}{row}"]
            if cell.value is None:
                continue
            if is_mom:
                cell.number_format = "+0.00%;-0.00%;0.00%"
                if isinstance(cell.value, (int, float)):
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=MOM_POS_BG if cell.value >= 0 else MOM_NEG_BG
                    )
            elif is_pct:
                cell.number_format = "0.00%"

    # 列幅（先頭 200 行でサンプリング）
    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1
    ):
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0) for c in col_cells
        )
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
    parser = argparse.ArgumentParser(description="月次レポート生成ツール")
    parser.add_argument("files",     nargs="*", help="対象 xlsx（省略時は自動スキャン）")
    parser.add_argument("--dir",     default=None,           help="スキャンディレクトリ")
    parser.add_argument("--output",  default=DEFAULT_OUTPUT, help="出力ファイル名")
    parser.add_argument("--ip-file", default=None,           help="SKU→IP マスタ")
    args = parser.parse_args()

    print("=== 月次レポート生成ツール ===")
    paths = collect_files(args.files, args.dir, args.output)
    print(f"対象: {len(paths)} ファイル")

    df_all       = load_all(paths)
    df_ip_master = load_ip_master(args.ip_file)

    # 当月・前月を自動判定
    valid_months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    if not valid_months:
        sys.exit("[エラー] 有効な日付データがありません。")
    curr_ym = valid_months[-1]
    prev_ym = valid_months[-2] if len(valid_months) >= 2 else None
    print(f"\n  当月: {curr_ym}  /  前月: {prev_ym or 'なし（先月比は N/A）'}")

    df_curr = df_all[df_all["__ym__"] == curr_ym]
    df_prev = df_all[df_all["__ym__"] == prev_ym] if prev_ym else None

    print("  集計中...")
    sheets = {
        "概要":  build_overview(df_curr, df_prev, curr_ym, prev_ym),
        "月別":  build_monthly(df_all),
        "店舗別": build_by_store(df_curr, df_prev),
        "商品別": build_by_product(df_curr, df_prev),
    }
    if df_ip_master is not None:
        # IP別は店舗別の後・商品別の前に挿入
        sheets = {k: v for k, v in list(sheets.items())[:3]} | \
                 {"IP別": build_by_ip(df_curr, df_prev, df_ip_master)} | \
                 {"商品別": sheets["商品別"]}
    else:
        print("  ⚠ IP マスタなし → IP別シートをスキップ")

    write_excel(args.output, sheets)


if __name__ == "__main__":
    main()
