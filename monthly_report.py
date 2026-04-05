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
import calendar
import re
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
MOM_POS_BG   = "E2EFDA"   # MoM positive (light green)
MOM_NEG_BG   = "FCE4D6"   # MoM negative (light red)
SECTION_BG   = "BDD7EE"   # Overview section divider

# ── 英語列名マッピング ─────────────────────────────────────────
_BASE_MAP: dict[str, str] = {
    # Dimensions
    "年月": "Month", "店舗": "Store", "IP名称": "IP",
    "商品コード": "SKU Code", "商品名": "SKU Name",
    "指標": "KPI", "先月比": "MoM%",
    # General KPIs
    "購入金額合計（税込）": "Sales",
    "トランザクション数":   "TXNs",
    "購入点数":            "Qty",
    "連帯率":              "UPT",
    "客単価":              "ATV",
    "金額構成比":          "Sales Mix%",
    "TXN構成比":           "TXN Mix%",
    "数量構成比":          "Qty Mix%",
    # Tax-free
    "免税購入金額合計（税込）": "TF Sales",
    "免税比率":               "TF %",
    # Member (from _agg_by / metrics_for)
    "会員金額（税込）":        "Mbr Sales",
    "会員金額（税込）構成比":   "Mbr Sales Mix%",
    "会員購入点数":            "Mbr Qty",
    "会員購入点数構成比":       "Mbr Qty Mix%",
    "会員人数":               "Active Mbr",
    "会員客単価":             "Mbr ATV",
    "会員連帯率":             "Mbr UPT",
    "会員購入頻度":           "Frequency",
    # Cross-tab short names
    "購入金額": "Sales", "TXN数": "TXNs", "点数": "Qty", "免税金額": "TF Sales",
    # Member sheet (_member_metrics)
    "TXN単価": "Mbr TXN ATV", "平均購入回数": "Avg Visits",
    # Budget
    "Month_JPY_MTD":          "Month MTD",
    "Month_Progress（％）":    "Month Progress%",
    "Month_Target_Reach（％）": "Month Target Reach%",
    "GAP（％）":               "GAP%",
}
_MOM_JP = "_先月比"
_MOM_EN = "_MoM%"

def _to_en(col: str) -> str:
    """列名1つを英語略称に変換（_先月比サフィックス・クロス集計列も処理）。"""
    if col in _BASE_MAP:
        return _BASE_MAP[col]
    if col.endswith(_MOM_JP):
        base = col[:-len(_MOM_JP)]
        if base in _BASE_MAP:
            return _BASE_MAP[base] + _MOM_EN
        # クロス集計の MoM列: "{store}_{metric}_先月比" → 再帰でベース変換
        return _to_en(base) + _MOM_EN
    # クロス集計列 "{store}_{metric}"（長い名前を優先してマッチ）
    for jp, en in sorted(_BASE_MAP.items(), key=lambda x: -len(x[0])):
        sfx = f"_{jp}"
        if col.endswith(sfx):
            return col[:-len(sfx)] + f"_{en}"
    return col

def _rename_df(df: pd.DataFrame) -> pd.DataFrame:
    """DataFrame の全列名を英語略称に変換して返す。"""
    return df.rename(columns={c: _to_en(c) for c in df.columns})

# 小数1桁（#,##0.0）で表示する英語列名セット
_FLOAT_EN = {"UPT", "ATV", "TF UPT", "TF ATV", "Mbr ATV", "Mbr UPT",
             "Frequency", "Avg Visits", "Mbr TXN ATV"}

def _is_float_col(col: str) -> bool:
    """列が #,##0.0 フォーマット対象かを返す（クロス集計列も含む）。"""
    if "[¥100M]" in col:
        return True
    if col in _FLOAT_EN:
        return True
    for fc in _FLOAT_EN:
        if col.endswith(f"_{fc}"):
            return True
    return False
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
    # ※ POS番号が空でも有効な取引（オンライン注文→店舗受取）があるため除外対象外
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


def load_budget(budget_file):
    """売上予算管理表（月別・店舗別）を読み込む。
    フォーマット:
      行0: タイトル / 行1: 注釈 / 行2: ヘッダー（店舗コード(H6 CODE) | 店舗名 | 1月…12月 | 年間合計 | 月平均）
      データ行、最終行に全店合計（除外）
    戻り値: {"store_col": str, "month_cols": {1:"1月",…}, "df": DataFrame}
    """
    if budget_file:
        path = Path(budget_file)
    else:
        keywords = ["budget", "target", "予算", "目標"]
        candidates = [p for p in sorted(Path(".").glob("*.xlsx"))
                      if any(k in p.stem.lower() for k in keywords)]
        if not candidates:
            return None
        path = candidates[0]

    print(f"  Budget ファイル: {path.name}")
    df = pd.read_excel(path, header=2, keep_default_na=False, engine="openpyxl")

    # 店舗コード列 = 最初の列
    store_col = df.columns[0]
    # 全店合計行・空行を除外
    df = df[~df[store_col].astype(str).str.contains("合計|total|grand", case=False, na=False)]
    df = df[df[store_col].astype(str).str.strip().replace("nan", "") != ""]
    df[store_col] = df[store_col].astype(str).str.strip()

    # 月別列を検出（"1月"〜"12月"）
    month_cols = {}
    for col in df.columns:
        m = re.match(r"^(\d{1,2})月$", str(col).strip())
        if m:
            month_cols[int(m.group(1))] = col

    print(f"    店舗数: {len(df)}  月別列: {sorted(month_cols.keys())}")
    return {"store_col": store_col, "month_cols": month_cols, "df": df}


def get_month_budget(budget_data, curr_ym):
    """当月の予算合計（全店）と店舗別予算 dict を返す。"""
    if budget_data is None:
        return None, {}
    month_num = int(curr_ym[5:7])
    col = budget_data["month_cols"].get(month_num)
    if col is None:
        return None, {}
    df        = budget_data["df"]
    store_col = budget_data["store_col"]
    amounts   = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return float(amounts.sum()), dict(zip(df[store_col].astype(str), amounts))


def calc_month_progress(df_curr, curr_ym):
    """データ内の最終日 ÷ 月の総日数 = 当月進捗率。"""
    valid  = df_curr[df_curr["__ym__"] != "日付不明"]
    max_dt = valid[COL_PICKUP_TIME].max()
    if pd.isna(max_dt):
        return None
    year, month = int(curr_ym[:4]), int(curr_ym[5:7])
    return max_dt.day / calendar.monthrange(year, month)[1]


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
    # 会員
    "会員金額（税込）",
    "会員金額（税込）構成比",
    "会員購入点数",
    "会員購入点数構成比",
    "会員人数",
    "会員客単価",
    "会員連帯率",
    "会員購入頻度",
]

def metrics_for(df) -> dict:
    txn    = df[COL_TXN].nunique()
    qty    = df[COL_QTY].sum()
    amount = df[COL_AMOUNT].sum()
    exempt = df.loc[df["__exempt__"], COL_AMOUNT].sum()

    # 会員
    df_m    = df[df[COL_MEMBER].astype(str).str.strip() != ""]
    m_count = df_m[COL_MEMBER].nunique()
    m_txn   = df_m[COL_TXN].nunique()
    m_qty   = df_m[COL_QTY].sum()
    m_amt   = df_m[COL_AMOUNT].sum()

    return {
        "購入金額合計（税込）":     amount,
        "トランザクション数":      txn,
        "購入点数":               qty,
        "連帯率":                round(qty    / txn,     2) if txn     else 0,
        "客単価":                round(amount / txn,     1) if txn     else 0,
        "免税購入金額合計（税込）": exempt,
        "免税比率":               exempt / amount            if amount  else 0,
        # 会員
        "会員金額（税込）":        m_amt,
        "会員金額（税込）構成比":   m_amt / amount            if amount  else 0,
        "会員購入点数":           m_qty,
        "会員購入点数構成比":      m_qty / qty               if qty     else 0,
        "会員人数":               m_count,
        "会員客単価":             round(m_amt / m_count, 1)  if m_count else 0,
        "会員連帯率":             round(m_qty / m_count, 2)  if m_count else 0,
        "会員購入頻度":           round(m_txn / m_count, 2)  if m_count else 0,
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

def build_overview(df_curr, df_prev, curr_ym, prev_ym, budget_data=None):
    """Sheet1: 縦型 KPI（KPI / 当月 / 前月 / MoM%）。金額は ¥100M 換算。
    budget_data が指定された場合は下部に予算比 KPI を追加する。
    """
    AMT_KEYS = {"購入金額合計（税込）", "免税購入金額合計（税込）", "会員金額（税込）"}

    curr = metrics_for(df_curr)
    prev = metrics_for(df_prev) if df_prev is not None else {}

    rows = []
    prev_col = prev_ym or "Prev Month"
    for key in METRIC_ORDER:
        en  = _BASE_MAP.get(key, key)
        c   = curr[key]
        p   = prev.get(key)
        m   = mom(c, p)
        if key in AMT_KEYS:
            label  = f"{en} [¥100M]"
            c_disp = round(c / OKU, 4)
            p_disp = round(p / OKU, 4) if p is not None else None
        else:
            label, c_disp, p_disp = en, c, p
        rows.append({"KPI": label, curr_ym: c_disp, prev_col: p_disp, "MoM%": m})

    # ── Budget KPI（Budget ファイルがある場合のみ）──────────
    if budget_data is not None:
        month_target, _ = get_month_budget(budget_data, curr_ym)
        month_progress  = calc_month_progress(df_curr, curr_ym)
        month_mtd       = df_curr[COL_AMOUNT].sum()

        if month_target and month_target > 0:
            target_reach = month_mtd / month_target
            gap          = month_mtd / month_target - 1
        else:
            target_reach = gap = None

        rows += [
            {"KPI": "── Budget ──────────────", curr_ym: None, prev_col: None, "MoM%": None},
            {"KPI": "Month Progress%",         curr_ym: month_progress, prev_col: None, "MoM%": None},
            {"KPI": "Month MTD",               curr_ym: month_mtd,      prev_col: None, "MoM%": None},
            {"KPI": "Month Target",            curr_ym: month_target,   prev_col: None, "MoM%": None},
            {"KPI": "Month Target Reach%",     curr_ym: target_reach,   prev_col: None, "MoM%": None},
            {"KPI": "GAP%",                    curr_ym: gap,            prev_col: None, "MoM%": None},
        ]

    return pd.DataFrame(rows)


def build_monthly(df_all, budget_data=None):
    """Sheet2: 月別推移（年月 | [予算列] | 各指標 | 先月比 … 交互配置）。"""
    months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")

    rows = []
    prev_m = {}
    for ym in months:
        df_ym  = df_all[df_all["__ym__"] == ym]
        curr_m = metrics_for(df_ym)
        row    = {"年月": ym}

        # ── 予算比 KPI ──────────────────────────────────────────
        if budget_data is not None:
            month_target, _ = get_month_budget(budget_data, ym)
            month_progress  = calc_month_progress(df_ym, ym)
            month_mtd       = df_ym[COL_AMOUNT].sum()
            if month_target and month_target > 0:
                reach = month_mtd / month_target
                gap   = month_mtd / month_target - 1
            else:
                month_target = reach = gap = None
            row["Month_Target"]           = month_target
            row["Month_Target_Reach（％）"] = reach
            row["GAP（％）"]               = gap

        for key in METRIC_ORDER:
            row[key]            = curr_m[key]
            row[f"{key}_先月比"] = mom(curr_m[key], prev_m.get(key))
        rows.append(row)
        prev_m = curr_m

    # 列順: 年月, [予算3列], (指標, 先月比) × 7
    ordered = ["年月"]
    if budget_data is not None:
        ordered += ["Month_Target", "Month_Target_Reach（％）", "GAP（％）"]
    for key in METRIC_ORDER:
        ordered += [key, f"{key}_先月比"]
    df = pd.DataFrame(rows)
    return df[[c for c in ordered if c in df.columns]]


def _agg_by(df, group_keys):
    """グループキーで基本集計し DataFrame を返す（会員 KPI 含む）。"""
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

    # 会員集計
    df_m  = df[df[COL_MEMBER].astype(str).str.strip() != ""]
    m_grp = df_m.groupby(group_keys, sort=False)
    m_agg = pd.DataFrame({
        "会員金額（税込）":  m_grp[COL_AMOUNT].sum(),
        "会員購入点数":     m_grp[COL_QTY].sum(),
        "会員人数":        m_grp[COL_MEMBER].nunique(),
        "_m_txn":         m_grp[COL_TXN].nunique(),
    }).reset_index().fillna(0)

    base = base.merge(m_agg, on=group_keys, how="left").fillna(0)
    nan_ = float("nan")
    base["会員金額（税込）構成比"] = (
        base["会員金額（税込）"] / base["購入金額合計（税込）"].replace(0, nan_)
    )
    base["会員購入点数構成比"] = (
        base["会員購入点数"] / base["購入点数"].replace(0, nan_)
    )
    base["会員客単価"]   = (base["会員金額（税込）"] / base["会員人数"].replace(0, nan_)).round(1)
    base["会員連帯率"]   = (base["会員購入点数"]    / base["会員人数"].replace(0, nan_)).round(2)
    base["会員購入頻度"] = (base["_m_txn"]          / base["会員人数"].replace(0, nan_)).round(2)
    return base.drop(columns=["_m_txn"])


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


def build_by_store(df_curr, df_prev, budget_data=None, curr_ym=None):
    """Sheet3: 店舗別。"""
    t_amt = df_curr[COL_AMOUNT].sum()
    t_txn = df_curr[COL_TXN].nunique()

    df = _agg_by(df_curr, [COL_STORE])
    df["金額構成比"] = df["購入金額合計（税込）"] / t_amt if t_amt else 0
    df["TXN構成比"]  = df["トランザクション数"]  / t_txn if t_txn else 0

    # ── 店舗別予算比 KPI ─────────────────────────────────────
    has_budget = budget_data is not None and curr_ym is not None
    if has_budget:
        _, store_budgets  = get_month_budget(budget_data, curr_ym)
        month_progress    = calc_month_progress(df_curr, curr_ym)

        def _lookup_budget(store_val):
            """店舗値から予算を検索する。
            1) 完全一致  2) [CODE]店名 形式からコード部分を抽出して照合
            """
            s = str(store_val).strip()
            # 完全一致
            if s in store_budgets:
                return store_budgets[s]
            # [CODE] プレフィックスを抽出（例: "[JP16]なんば店" → "JP16"）
            m = re.match(r"^\[([^\]]+)\]", s)
            if m:
                code = m.group(1).strip()
                if code in store_budgets:
                    return store_budgets[code]
            # 予算ファイルのキーに店舗値が部分一致するか逆引き
            for bkey, bval in store_budgets.items():
                if bkey in s or s in bkey:
                    return bval
            return None

        targets, reaches, gaps = [], [], []
        for store in df[COL_STORE]:
            target = _lookup_budget(store)
            mtd    = float(df.loc[df[COL_STORE] == store,
                                  "購入金額合計（税込）"].values[0])
            if target and target > 0:
                reach = mtd / target
                gap   = mtd / target - 1
            else:
                target = reach = gap = None
            targets.append(target)
            reaches.append(reach)
            gaps.append(gap)
        df["Month_Target"]           = targets
        df["Month_Target_Reach（％）"] = reaches
        df["GAP（％）"]               = gaps

    df = _add_mom_cols(df, df_prev, [COL_STORE],
                       ["購入金額合計（税込）", "トランザクション数",
                        "購入点数", "免税購入金額合計（税込）",
                        "会員金額（税込）", "会員購入点数", "会員人数",
                        "会員客単価", "会員連帯率", "会員購入頻度"])

    col_order = [COL_STORE]
    if has_budget:
        col_order += ["Month_Target", "Month_Target_Reach（％）", "GAP（％）"]
    col_order += [
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "TXN構成比",  "トランザクション数_先月比",
        "購入点数", "連帯率", "客単価",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
        # 会員
        "会員金額（税込）", "会員金額（税込）構成比", "会員金額（税込）_先月比",
        "会員購入点数", "会員購入点数構成比", "会員購入点数_先月比",
        "会員人数", "会員人数_先月比",
        "会員客単価", "会員客単価_先月比",
        "会員連帯率", "会員連帯率_先月比",
        "会員購入頻度", "会員購入頻度_先月比",
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
                        "購入点数", "免税購入金額合計（税込）",
                        "会員金額（税込）", "会員購入点数", "会員人数",
                        "会員客単価", "会員連帯率", "会員購入頻度"])

    col_order = [
        COL_IP,
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "トランザクション数_先月比",
        "購入点数", "数量構成比", "購入点数_先月比",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
        # 会員
        "会員金額（税込）", "会員金額（税込）構成比", "会員金額（税込）_先月比",
        "会員購入点数", "会員購入点数構成比", "会員購入点数_先月比",
        "会員人数", "会員人数_先月比",
        "会員客単価", "会員客単価_先月比",
        "会員連帯率", "会員連帯率_先月比",
        "会員購入頻度", "会員購入頻度_先月比",
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
                        "購入点数", "免税購入金額合計（税込）",
                        "会員金額（税込）", "会員購入点数", "会員人数",
                        "会員客単価", "会員連帯率", "会員購入頻度"])

    col_order = [
        COL_PROD_CODE, COL_PRODUCT,
        "購入金額合計（税込）", "金額構成比", "購入金額合計（税込）_先月比",
        "トランザクション数",  "トランザクション数_先月比",
        "購入点数", "数量構成比", "購入点数_先月比",
        "免税購入金額合計（税込）", "免税比率", "免税購入金額合計（税込）_先月比",
        # 会員
        "会員金額（税込）", "会員金額（税込）構成比", "会員金額（税込）_先月比",
        "会員購入点数", "会員購入点数構成比", "会員購入点数_先月比",
        "会員人数", "会員人数_先月比",
        "会員客単価", "会員客単価_先月比",
        "会員連帯率", "会員連帯率_先月比",
        "会員購入頻度", "会員購入頻度_先月比",
    ]
    df = df[[c for c in col_order if c in df.columns]]
    return df.sort_values("購入金額合計（税込）", ascending=False).reset_index(drop=True)


# ── 短縮メトリクス名（クロス集計用）─────────────────────────
# 店舗名プレフィックスが付くため列名を短縮する
_CROSS_METRICS_ORDER = [
    "購入金額", "金額構成比", "購入金額_先月比",
    "TXN数",   "TXN数_先月比",
    "点数",    "数量構成比",  "点数_先月比",
    "免税金額", "免税比率",   "免税金額_先月比",
]


def _agg_cross(df, all_keys):
    """クロス集計用の短縮名集計（行キー + COL_STORE）。"""
    grp    = df.groupby(all_keys, sort=False)
    ex_grp = df[df["__exempt__"]].groupby(all_keys, sort=False)
    base = pd.DataFrame({
        "購入金額": grp[COL_AMOUNT].sum(),
        "TXN数":   grp[COL_TXN].nunique(),
        "点数":    grp[COL_QTY].sum(),
        "免税金額": ex_grp[COL_AMOUNT].sum(),
    }).reset_index().fillna(0)
    return base


def _build_cross_tab(df_curr, df_prev, row_keys):
    """
    縦 = row_keys / 横 = 店舗 のクロス集計ワイドテーブルを返す。
    各店舗につき: 購入金額 | 金額構成比 | 先月比 | TXN数 | 先月比 |
                  点数 | 数量構成比 | 先月比 | 免税金額 | 免税比率 | 先月比
    """
    all_keys = row_keys + [COL_STORE]

    curr = _agg_cross(df_curr, all_keys)
    prev = _agg_cross(df_prev, all_keys) if df_prev is not None else None

    # 先月比（店舗×行キーレベルで計算）
    has_mom = prev is not None
    if has_mom:
        merged = curr.merge(prev, on=all_keys, how="left", suffixes=("", "_前月"))
        for m in ["購入金額", "TXN数", "点数", "免税金額"]:
            merged[f"{m}_先月比"] = merged.apply(
                lambda r, c=m: mom(r[c], r.get(f"{c}_前月")), axis=1
            )
        curr = merged.drop(columns=[c for c in merged.columns if c.endswith("_前月")])

    # 店舗別合計（構成比の分母）
    store_tot = curr.groupby(COL_STORE)[["購入金額", "点数"]].sum().rename(
        columns={"購入金額": "_t_amt", "点数": "_t_qty"}
    )
    curr = curr.merge(store_tot, on=COL_STORE, how="left")
    curr["金額構成比"] = curr["購入金額"] / curr["_t_amt"].replace(0, float("nan"))
    curr["数量構成比"] = curr["点数"]     / curr["_t_qty"].replace(0, float("nan"))
    curr["免税比率"]   = curr["免税金額"] / curr["購入金額"].replace(0, float("nan"))
    curr = curr.drop(columns=["_t_amt", "_t_qty"])

    # 使用するメトリクス列
    m_cols = [m for m in _CROSS_METRICS_ORDER if m in curr.columns]

    # ワイド形式に変換：店舗ごとにメトリクスを横展開
    stores  = sorted(curr[COL_STORE].unique())
    wide    = curr[row_keys].drop_duplicates().copy()
    for store in stores:
        sub = curr[curr[COL_STORE] == store][row_keys + m_cols].copy()
        sub = sub.rename(columns={m: f"{store}_{m}" for m in m_cols})
        wide = wide.merge(sub, on=row_keys, how="left")

    # 全店舗合計で降順ソート
    amt_cols  = [f"{s}_購入金額" for s in stores if f"{s}_購入金額" in wide.columns]
    wide["_total"] = wide[amt_cols].sum(axis=1)
    wide = wide.sort_values("_total", ascending=False).drop(columns=["_total"])

    return wide.fillna(0).reset_index(drop=True)


def build_ip_x_store(df_curr, df_prev, df_ip_master):
    """Sheet6: IP × 店舗 クロス集計。"""
    def add_ip(df):
        d = df.copy()
        d[COL_PROD_CODE] = d[COL_PROD_CODE].astype(str).str.strip()
        d = d.merge(df_ip_master[[COL_PROD_CODE, COL_IP]], on=COL_PROD_CODE, how="left")
        d[COL_IP] = d[COL_IP].fillna("IP未設定")
        return d

    return _build_cross_tab(add_ip(df_curr),
                             add_ip(df_prev) if df_prev is not None else None,
                             row_keys=[COL_IP])


def build_product_x_store(df_curr, df_prev):
    """Sheet7: 商品 × 店舗 クロス集計。"""
    return _build_cross_tab(df_curr, df_prev,
                             row_keys=[COL_PROD_CODE, COL_PRODUCT])


# ══════════════════════════════════════════════════════════════
#  会員関連シート
# ══════════════════════════════════════════════════════════════

def _member_metrics(df_scope, df_total):
    """
    df_scope: 集計対象（例: ある店舗のデータ）
    df_total: 構成比の分母（同スコープの全顧客）
    """
    df_m   = df_scope[df_scope[COL_MEMBER].astype(str).str.strip() != ""]
    t_amt  = df_total[COL_AMOUNT].sum()
    t_qty  = df_total[COL_QTY].sum()

    m_count = df_m[COL_MEMBER].nunique()
    m_txn   = df_m[COL_TXN].nunique()
    m_qty   = df_m[COL_QTY].sum()
    m_amt   = df_m[COL_AMOUNT].sum()

    return {
        "会員人数":     m_count,
        "購入金額":     m_amt,
        "金額構成比":   m_amt / t_amt       if t_amt    else 0,
        "TXN数":       m_txn,
        "点数":        m_qty,
        "数量構成比":   m_qty / t_qty       if t_qty    else 0,
        "会員客単価":   round(m_amt / m_count, 1) if m_count else 0,
        "TXN単価":     round(m_amt / m_txn,   1) if m_txn   else 0,
        "平均購入回数":  round(m_txn / m_count, 2) if m_count else 0,
    }


def build_member_sheet(df_curr, df_prev):
    """Sheet8: 会員 全体サマリ（1行 + 先月比）。"""
    curr = _member_metrics(df_curr, df_curr)
    prev = _member_metrics(df_prev, df_prev) if df_prev is not None else {}

    row = {}
    MOM_TARGETS = {"購入金額", "TXN数", "点数"}
    col_order   = []

    for key in ["会員人数", "購入金額", "金額構成比",
                "TXN数", "点数", "数量構成比",
                "会員客単価", "TXN単価", "平均購入回数"]:
        row[key] = curr[key]
        col_order.append(key)
        if key in MOM_TARGETS:
            mom_col = f"{key}_先月比"
            row[mom_col] = mom(curr[key], prev.get(key))
            col_order.append(mom_col)

    return pd.DataFrame([row])[col_order]


def build_store_x_member(df_curr, df_prev):
    """Sheet9: 店舗×会員 — 店舗ごとの会員集計。"""
    stores = sorted(df_curr[COL_STORE].unique())
    MOM_TARGETS = {"購入金額", "TXN数", "点数"}
    rows = []

    for store in stores:
        dc = df_curr[df_curr[COL_STORE] == store]
        dp = (df_prev[df_prev[COL_STORE] == store]
              if df_prev is not None else None)
        dp = dp if dp is not None and len(dp) > 0 else None

        curr = _member_metrics(dc, dc)
        prev = _member_metrics(dp, dp) if dp is not None else {}

        row = {COL_STORE: store}
        for key in ["会員人数", "購入金額", "金額構成比",
                    "TXN数", "点数", "数量構成比",
                    "会員客単価", "TXN単価", "平均購入回数"]:
            row[key] = curr[key]
            if key in MOM_TARGETS:
                row[f"{key}_先月比"] = mom(curr[key], prev.get(key))
        rows.append(row)

    col_order = [COL_STORE,
                 "会員人数",
                 "購入金額", "金額構成比", "購入金額_先月比",
                 "TXN数",   "TXN数_先月比",
                 "点数",    "数量構成比",  "点数_先月比",
                 "会員客単価", "TXN単価", "平均購入回数"]

    df = pd.DataFrame(rows)
    df = df[[c for c in col_order if c in df.columns]]
    return df.sort_values("購入金額", ascending=False).reset_index(drop=True)


# ══════════════════════════════════════════════════════════════
#  説明シート（Glossary）
# ══════════════════════════════════════════════════════════════

_GLO_COLS = ["日本語指標", "英語翻訳 (Full)", "英語略称", "英語計算式 (Formula)", "日本語計算式"]

def _sec(text):
    return {"日本語指標": text, "英語翻訳 (Full)": "",
            "英語略称": "", "英語計算式 (Formula)": "", "日本語計算式": ""}

def _kpi(jp, full, en, formula, jp_formula):
    return {"日本語指標": jp, "英語翻訳 (Full)": full, "英語略称": en,
            "英語計算式 (Formula)": formula, "日本語計算式": jp_formula}

def build_glossary():
    """Glossary シート: KPI 定義一覧（3セクション）。"""
    rows = [
        _sec("【基本指標】  売上を構成する要素を細かく分解したものです。"
             "  全体売上 ＝ TXN数 × 客単価（連帯率 × 平均単価）"),
        _kpi("全体購入金額合計", "Total Gross Sales",         "Sales",
             "Σ(Tax-Incl. Amount)",      "全取引の税込販売金額の合計"),
        _kpi("TXN数",           "Number of Transactions",    "TXNs",
             "COUNTD(TXN_KEY)",          "取引件数（レジを通った回数）"),
        _kpi("点数",            "Units Sold",                "Qty",
             "Σ(Quantity)",              "販売済みの商品の総個数"),
        _kpi("連帯率",          "Units Per Transaction",     "UPT",
             "Qty ÷ TXNs",              "点数/TXN数（1回あたりの買上げ点数）"),
        _kpi("客単価",          "Average Transaction Value", "ATV",
             "Sales ÷ TXNs",            "金額/TXN数（1回あたりの買上げ金額）"),

        _sec("【免税（Tax Free）指標】  インバウンド需要を測る重要な指標です。"),
        _kpi("免税購入金額合計", "Total Tax-Free Sales",              "TF Sales",
             "Σ(Amount | Tax = 0)",          "免税取引の税込販売金額の合計"),
        _kpi("免税比率",        "Tax-Free Sales Ratio",              "TF %",
             "TF Sales ÷ Sales",              "免税金額/全体金額×100%"),
        _kpi("免税TXN数",       "Tax-Free Transactions",             "TF TXNs",
             "COUNTD(TXN_KEY | Tax = 0)",    "免税取引の件数"),
        _kpi("免税数量",        "Tax-Free Units Sold",               "TF Qty",
             "Σ(Qty | Tax = 0)",             "免税で販売した商品の総個数"),
        _kpi("免税連帯率",      "Tax-Free Units Per Transaction",    "TF UPT",
             "TF Qty ÷ TF TXNs",             "免税数量/免税TXN数"),
        _kpi("免税客単価",      "Average TF Transaction Value",      "TF ATV",
             "TF Sales ÷ TF TXNs",           "免税金額/免税TXN数"),

        _sec("【会員（CRM）指標】  リピーター戦略やファン化を測る指標です。"),
        _kpi("会員金額",        "Total Member Sales",                "Mbr Sales",
             "Σ(Amount | Member ID ≠ blank)",        "会員証が提示された取引の合計金額"),
        _kpi("会員購入点数",    "Member Units Sold",                 "Mbr Qty",
             "Σ(Qty | Member ID ≠ blank)",           "会員が購入した商品の総個数"),
        _kpi("会員人数",        "Number of Unique Members",          "Active Mbr",
             "COUNTD(Member ID | Member ID ≠ blank)", "期間中に購入のあったユニークな会員数"),
        _kpi("会員客単価",      "Member Average Transaction Value",  "Mbr ATV",
             "Mbr Sales ÷ Active Mbr",               "会員金額/会員TXN数"),
        _kpi("会員連帯率",      "Member Units Per Transaction",      "Mbr UPT",
             "Mbr Qty ÷ Active Mbr",                 "会員点数/会員TXN数"),
        _kpi("会員購入頻度",    "Purchase Frequency",                "Frequency",
             "Mbr TXNs ÷ Active Mbr",                "会員TXN数/会員人数（期間中の来店回数）"),
    ]
    return pd.DataFrame(rows, columns=_GLO_COLS)


def _style_glossary(ws):
    """Glossary シート専用スタイル（5列: 日本語 / Full翻訳 / 略称 / 計算式EN / 計算式JP）。"""
    h_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    s_fill  = PatternFill("solid", fgColor=HEADER_BG)
    k_alt   = PatternFill("solid", fgColor="EBF3FB")
    n_font  = Font(name="Arial", size=10)

    # ヘッダー行
    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

    # セクション行の判定: 英語略称列（B = col 3、Full翻訳が追加されたので）が空
    # 5列構成: A=日本語 B=Full C=略称 D=計算式EN E=計算式JP → section check: col 3 (C) empty
    kpi_count = 0
    for row in range(2, ws.max_row + 1):
        is_section = str(ws.cell(row=row, column=3).value or "").strip() == ""
        for col in range(1, 6):
            cell = ws.cell(row=row, column=col)
            if is_section:
                cell.font      = s_font
                cell.fill      = s_fill
                cell.alignment = Alignment(horizontal="left", vertical="center",
                                           wrap_text=True)
            else:
                cell.font = n_font
                kpi_count += 1
                if kpi_count % 2 == 0:
                    cell.fill = k_alt
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        # セクション行は A:E をマージ
        if is_section:
            ws.merge_cells(f"A{row}:E{row}")

    # 列幅固定
    ws.column_dimensions["A"].width = 18   # 日本語指標
    ws.column_dimensions["B"].width = 34   # 英語翻訳 (Full)
    ws.column_dimensions["C"].width = 14   # 英語略称
    ws.column_dimensions["D"].width = 36   # 英語計算式
    ws.column_dimensions["E"].width = 36   # 日本語計算式
    ws.row_dimensions[1].height     = 20


# ══════════════════════════════════════════════════════════════
#  スタイル & 出力
# ══════════════════════════════════════════════════════════════

def _col_letters(df):
    """列名 → Excelの列レター の dict を返す。"""
    return {col: get_column_letter(i + 1) for i, col in enumerate(df.columns)}


def style_sheet(ws, df):
    """表頭スタイル・列幅・フォーマット・MoM%色分けを適用。フォント: Arial。"""
    h_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

    letters = _col_letters(df)
    n_font  = Font(name="Arial", size=10)

    for col, letter in letters.items():
        is_mom    = "_MoM%" in col or col == "MoM%"
        is_pct    = not is_mom and "%" in col
        is_target = "Target" in col and "%" not in col
        is_float  = not is_mom and not is_pct and _is_float_col(col)

        for row in range(2, ws.max_row + 1):
            cell = ws[f"{letter}{row}"]
            if cell.value is None:
                continue
            cell.font = n_font
            if is_mom:
                cell.number_format = "+0.00%;-0.00%;0.00%"
                if isinstance(cell.value, (int, float)):
                    cell.fill = PatternFill(
                        "solid",
                        fgColor=MOM_POS_BG if cell.value >= 0 else MOM_NEG_BG
                    )
            elif is_pct:
                cell.number_format = "0.00%"
            elif is_target:
                cell.number_format = "#,##0"
            elif is_float:
                cell.number_format = "#,##0.0"
            elif isinstance(cell.value, (int, float)):
                cell.number_format = "#,##0"

    # 列幅（先頭 200 行でサンプリング）
    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1
    ):
        max_len = max(
            (len(str(c.value)) if c.value is not None else 0) for c in col_cells
        )
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"


_OVW_FLOAT_LABELS = {"UPT", "ATV", "TF UPT", "TF ATV", "Mbr ATV", "Mbr UPT",
                     "Frequency", "Avg Visits"}
_OVW_INT_LABELS   = {"Month MTD", "Month Target", "TXNs", "Mbr Qty",
                     "Qty", "Active Mbr", "TF TXNs", "TF Qty"}

def _style_overview(ws):
    """Overview シート専用: 英語行ラベルを見てセル単位でフォーマット＋Arial 適用。"""
    n_font = Font(name="Arial", size=10)
    for row in range(2, ws.max_row + 1):
        label = str(ws.cell(row=row, column=1).value or "")
        for col in range(2, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            cell.font = n_font
            if cell.value is None or not isinstance(cell.value, (int, float)):
                continue
            if "%" in label:                           # PCT rows
                cell.number_format = "0.0%"
            elif "[¥100M]" in label:                   # Billion-yen rows
                cell.number_format = "#,##0.0"
            elif label in _OVW_INT_LABELS:             # Integer KPIs
                cell.number_format = "#,##0"
            elif label in _OVW_FLOAT_LABELS:           # Float KPIs (UPT/ATV…)
                cell.number_format = "#,##0.0"
            elif label in {"Month MTD", "Month Target"}:
                cell.number_format = "#,##0"
            elif ws.cell(row=1, column=col).value == "MoM%":
                cell.number_format = "+0.0%;-0.0%;0.0%"
                cell.fill = PatternFill("solid",
                    fgColor=MOM_POS_BG if cell.value >= 0 else MOM_NEG_BG)
            else:
                cell.number_format = "#,##0"


def write_excel(output_path, sheets: dict):
    print(f"\n  書込中: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            # Glossary はそのまま（既に英語列名）、他は列名を英語に変換
            out_df = df if name == "Glossary" else _rename_df(df)
            out_df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            if name == "Overview":
                style_sheet(ws, out_df)
                _style_overview(ws)
            elif name == "Glossary":
                _style_glossary(ws)
            else:
                style_sheet(ws, out_df)
            print(f"    [{name}] {len(out_df):,} 行")
    print(f"\n  完成！→ {output_path}")


# ══════════════════════════════════════════════════════════════
#  エントリポイント
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="月次レポート生成ツール")
    parser.add_argument("files",     nargs="*", help="対象 xlsx（省略時は自動スキャン）")
    parser.add_argument("--dir",     default=None,           help="スキャンディレクトリ")
    parser.add_argument("--output",  default=DEFAULT_OUTPUT, help="出力ファイル名")
    parser.add_argument("--ip-file",     default=None, help="SKU→IP マスタ")
    parser.add_argument("--budget-file", default=None, help="売上予算管理表 xlsx（省略時は自動検出）")
    args = parser.parse_args()

    print("=== 月次レポート生成ツール ===")
    paths = collect_files(args.files, args.dir, args.output)
    print(f"対象: {len(paths)} ファイル")

    df_all       = load_all(paths)
    df_ip_master = load_ip_master(args.ip_file)
    budget_data  = load_budget(args.budget_file)

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
        "Glossary":  build_glossary(),
        "Overview":  build_overview(df_curr, df_prev, curr_ym, prev_ym, budget_data),
        "Monthly":   build_monthly(df_all, budget_data),
        "By Store":  build_by_store(df_curr, df_prev, budget_data, curr_ym),
        "By SKU":    build_by_product(df_curr, df_prev),
    }
    if df_ip_master is not None:
        sheets = {k: v for k, v in list(sheets.items())[:4]} | \
                 {"By IP":      build_by_ip(df_curr, df_prev, df_ip_master)} | \
                 {"By SKU":     sheets["By SKU"]} | \
                 {"IP x Store": build_ip_x_store(df_curr, df_prev, df_ip_master)} | \
                 {"SKU x Store": build_product_x_store(df_curr, df_prev)}
    else:
        print("  ⚠ IP マスタなし → By IP / IP x Store シートをスキップ")
        sheets["SKU x Store"] = build_product_x_store(df_curr, df_prev)

    sheets["Member"]        = build_member_sheet(df_curr, df_prev)
    sheets["Store x Member"] = build_store_x_member(df_curr, df_prev)

    write_excel(args.output, sheets)


if __name__ == "__main__":
    main()
