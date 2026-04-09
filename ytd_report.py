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
MOM_POS_BG = "E2EFDA"
MOM_NEG_BG = "FCE4D6"

# ── 英語列名マッピング ─────────────────────────────────────────
_BASE_MAP: dict[str, str] = {
    # Dimensions
    "年月": "Month", "店舗": "Store", "IP名称": "IP",
    "商品コード": "SKU Code", "商品名": "SKU Name",
    "指標": "KPI",
    # General KPIs
    "全体購入金額合計（税込）": "Sales",
    "TXN数":                  "TXNs",
    "点数":                   "Qty",
    "連帯率":                 "UPT",
    "客単価":                 "ATV",
    # Tax-free
    "免税購入金額合計（税込）": "TF Sales",
    "免税比率":               "TF %",
    "免税TXN数":              "TF TXNs",
    "免税数量":               "TF Qty",
    "免税連帯率":              "TF UPT",
    "免税客単価":              "TF ATV",
    # Member
    "会員金額（税込）":        "Mbr Sales",
    "会員購入点数":            "Mbr Qty",
    "会員人数":               "Active Mbr",
    "会員客単価":             "Mbr ATV",
    "会員連帯率":             "Mbr UPT",
    "会員購入頻度":           "Frequency",
    # Member – store×month sheet
    "会員TXN数":              "Mbr TXNs",
    "会員購入金額Mix％":       "Mbr Sales Mix%",
    "会員人数Mix％":           "Mbr Count Mix%",
    # Budget – monthly
    "Month_Target":           "Month Target",
    "Month_Target_Reach（％）": "Month Target Reach%",
    "GAP（％）":               "GAP%",
    # Budget – yearly
    "Year_Target":             "Year Target",
    "Year_JPY_YTD":            "Year YTD",
    "Year_Time_Progress（％）": "Year Progress%",
    "YEAR_TARGET_REACH（％）":  "Year Target Reach%",
}
_MOM_JP = "_先月比"
_MOM_EN = "_MoM%"

def _to_en(col: str) -> str:
    if col in _BASE_MAP:
        return _BASE_MAP[col]
    if col.endswith(_MOM_JP):
        base = col[:-len(_MOM_JP)]
        if base in _BASE_MAP:
            return _BASE_MAP[base] + _MOM_EN
        return _to_en(base) + _MOM_EN
    for jp, en in sorted(_BASE_MAP.items(), key=lambda x: -len(x[0])):
        if col.endswith(f"_{jp}"):
            return col[:-len(jp)-1] + f"_{en}"
    return col

def _rename_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: _to_en(c) for c in df.columns})

_FLOAT_EN = {"UPT", "ATV", "TF UPT", "TF ATV", "Mbr ATV", "Mbr UPT", "Frequency"}

def _is_float_col(col: str) -> bool:
    if "[¥100M]" in col:
        return True
    if col in _FLOAT_EN:
        return True
    for fc in _FLOAT_EN:
        if col.endswith(f"_{fc}"):
            return True
    return False
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
    "会員購入点数",        # ← 追加
    "会員人数",
    "会員客単価",
    "会員連帯率",
    "会員購入頻度",
]

PCT_KPI  = {"免税比率"}
# 小数点1桁表示（千切りあり）
FLOAT_KPI = {"連帯率", "客単価", "免税連帯率", "免税客単価",
             "会員客単価", "会員連帯率", "会員購入頻度"}
OKU_KPI  = {"全体購入金額合計（税込）",
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


def load_budget(budget_file):
    """売上予算管理表（月別・店舗別）を読み込む。"""
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

    store_col = df.columns[0]
    df = df[~df[store_col].astype(str).str.contains("合計|total|grand", case=False, na=False)]
    df = df[df[store_col].astype(str).str.strip().replace("nan", "") != ""]
    df[store_col] = df[store_col].astype(str).str.strip()

    month_cols = {}
    for col in df.columns:
        m = re.match(r"^(\d{1,2})月$", str(col).strip())
        if m:
            month_cols[int(m.group(1))] = col

    print(f"    店舗数: {len(df)}  月別列: {sorted(month_cols.keys())}")
    return {"store_col": store_col, "month_cols": month_cols, "df": df}


def get_year_budget(budget_data):
    """全月・全店舗の年間合計予算を返す。"""
    if budget_data is None:
        return None
    total = sum(
        pd.to_numeric(budget_data["df"][col], errors="coerce").fillna(0).sum()
        for col in budget_data["month_cols"].values()
    )
    return float(total) if total > 0 else None


def calc_year_progress(df_all):
    """データの最終日 ÷ その年の総日数 = 年間進捗率。"""
    valid  = df_all[df_all["__ym__"] != "日付不明"]
    max_dt = valid[COL_PICKUP_TIME].max()
    if pd.isna(max_dt):
        return None
    year       = max_dt.year
    total_days = 366 if calendar.isleap(year) else 365
    elapsed    = (max_dt - pd.Timestamp(year, 1, 1)).days + 1
    return elapsed / total_days


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
        "会員購入点数":          m_qty,
        "会員人数":              m_count,
        "会員客単価":            round(m_amt  / m_count, 1) if m_count else 0,
        "会員連帯率":            round(m_qty  / m_count, 2) if m_count else 0,
        "会員購入頻度":          round(m_txn  / m_count, 2) if m_count else 0,
    }


# ══════════════════════════════════════════════════════════════
#  Sheet ビルド
# ══════════════════════════════════════════════════════════════

def build_overview(df_all, budget_data=None):
    """Sheet1: YTD 全体 KPI（縦型）。金額 KPI は ¥100M 換算。"""
    kpi = calc_kpi(df_all)
    rows = []
    for key in KPI_ORDER:
        val = kpi[key]
        en  = _BASE_MAP.get(key, key)
        if key in OKU_KPI:
            rows.append({"KPI": f"{en} [¥100M]", "YTD": round(val / OKU, 4)})
        else:
            rows.append({"KPI": en, "YTD": val})

    # ── Year Budget KPI ────────────────────────────────────────
    if budget_data is not None:
        year_target   = get_year_budget(budget_data)
        year_progress = calc_year_progress(df_all)
        year_ytd      = df_all[COL_AMOUNT].sum()

        if year_target and year_target > 0:
            reach = year_ytd / year_target
            gap   = year_ytd / year_target - 1
        else:
            year_target = reach = gap = None

        rows += [
            {"KPI": "── Year Budget ─────────────", "YTD": None},
            {"KPI": "Year Progress%",              "YTD": year_progress},
            {"KPI": "Year YTD",                    "YTD": year_ytd},
            {"KPI": "Year Target",                 "YTD": year_target},
            {"KPI": "Year Target Reach%",          "YTD": reach},
            {"KPI": "GAP%",                        "YTD": gap},
        ]

    return pd.DataFrame(rows)


def _get_month_budget_total(budget_data, ym):
    """当月の予算合計（全店）を返す。"""
    if budget_data is None:
        return None
    col = budget_data["month_cols"].get(int(ym[5:7]))
    if col is None:
        return None
    return float(pd.to_numeric(budget_data["df"][col], errors="coerce").fillna(0).sum())


def _get_store_year_budget(budget_data):
    """店舗ごとの年間合計予算を dict で返す。"""
    if budget_data is None:
        return {}
    df        = budget_data["df"]
    store_col = budget_data["store_col"]
    result = {}
    for _, row in df.iterrows():
        store = str(row[store_col]).strip()
        total = sum(float(pd.to_numeric(row.get(col, 0), errors="coerce") or 0)
                    for col in budget_data["month_cols"].values())
        result[store] = total
    return result


def _lookup_store_budget(store_val, store_budgets):
    """[CODE]店名 形式に対応した 3 段階の予算照合。"""
    s = str(store_val).strip()
    if s in store_budgets:
        return store_budgets[s]
    m = re.match(r"^\[([^\]]+)\]", s)
    if m:
        code = m.group(1).strip()
        if code in store_budgets:
            return store_budgets[code]
    for bkey, bval in store_budgets.items():
        if bkey in s or s in bkey:
            return bval
    return None


def _budget_kpis(ytd_amt, target):
    """(reach, gap) を返す。target が None/0 の場合は (None, None)。"""
    if target and target > 0:
        return ytd_amt / target, ytd_amt / target - 1
    return None, None


def build_monthly(df_all, budget_data=None):
    """Sheet2: 月別 KPI（年月 | [予算列] | KPI列…）。"""
    months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    rows = []
    for ym in months:
        df_ym = df_all[df_all["__ym__"] == ym]
        kpi   = calc_kpi(df_ym)
        row   = {"年月": ym}

        if budget_data is not None:
            target       = _get_month_budget_total(budget_data, ym)
            reach, gap   = _budget_kpis(df_ym[COL_AMOUNT].sum(), target)
            row["Month_Target"]           = target
            row["Month_Target_Reach（％）"] = reach
            row["GAP（％）"]               = gap

        row.update({k: kpi[k] for k in KPI_ORDER})
        rows.append(row)
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


def build_by_store(df_all, budget_data=None):
    """Sheet3: 店舗別 YTD KPI。"""
    df = _build_by_dim(df_all, [COL_STORE])

    if budget_data is not None:
        store_budgets = _get_store_year_budget(budget_data)
        targets, reaches, gaps = [], [], []
        for store in df[COL_STORE]:
            target     = _lookup_store_budget(store, store_budgets)
            ytd        = float(df.loc[df[COL_STORE] == store,
                                      "全体購入金額合計（税込）"].values[0])
            reach, gap = _budget_kpis(ytd, target)
            targets.append(target); reaches.append(reach); gaps.append(gap)
        df.insert(1, "Year_Target",            targets)
        df.insert(2, "YEAR_TARGET_REACH（％）", reaches)
        df.insert(3, "GAP（％）",              gaps)

    return df


def build_store_monthly(df_all):
    """店舗×月別 KPI シート。
    列: 店舗 | 年月 | 全体売上金額 | TXN数 | 会員購入金額 | 会員TXN数 |
        会員人数 | 会員購入金額Mix％ | 会員人数Mix％ | 客単価 | 会員客単価 |
        連帯率 | 会員連帯率
    会員人数Mix％ = 会員人数 ÷ TXN数（全体）
    """
    months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    stores = sorted(df_all[COL_STORE].dropna().unique())

    rows = []
    for store in stores:
        df_s = df_all[df_all[COL_STORE] == store]
        for ym in months:
            df_sm = df_s[df_s["__ym__"] == ym]
            if df_sm.empty:
                continue

            # ── 全体 ─────────────────────────────────────────
            txn = df_sm[COL_TXN].nunique()
            qty = df_sm[COL_QTY].sum()
            amt = df_sm[COL_AMOUNT].sum()

            # ── 会員 ─────────────────────────────────────────
            df_m    = df_sm[df_sm[COL_MEMBER].astype(str).str.strip() != ""]
            m_count = df_m[COL_MEMBER].nunique()
            m_txn   = df_m[COL_TXN].nunique()
            m_qty   = df_m[COL_QTY].sum()
            m_amt   = df_m[COL_AMOUNT].sum()

            rows.append({
                COL_STORE:          store,
                "年月":              ym,
                "全体購入金額合計（税込）": amt,
                "TXN数":             txn,
                "会員金額（税込）":    m_amt,
                "会員TXN数":          m_txn,
                "会員人数":           m_count,
                "会員購入金額Mix％":   m_amt / amt     if amt     else 0,
                "会員人数Mix％":       m_count / txn   if txn     else 0,
                "客単価":             round(amt / txn,     1) if txn     else 0,
                "会員客単価":         round(m_amt / m_count, 1) if m_count else 0,
                "連帯率":             round(qty / txn,     2) if txn     else 0,
                "会員連帯率":         round(m_qty / m_count, 2) if m_count else 0,
            })

    return pd.DataFrame(rows)


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
#  Glossary シート
# ══════════════════════════════════════════════════════════════

_GLO_COLS = ["日本語指標", "英語翻訳 (Full)", "英語略称", "英語計算式 (Formula)", "日本語計算式"]

def _sec(text):
    return {"日本語指標": text, "英語翻訳 (Full)": "",
            "英語略称": "", "英語計算式 (Formula)": "", "日本語計算式": ""}

def _kpi(jp, full, en, formula, jp_formula):
    return {"日本語指標": jp, "英語翻訳 (Full)": full, "英語略称": en,
            "英語計算式 (Formula)": formula, "日本語計算式": jp_formula}

def build_glossary():
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
             "Mbr Sales ÷ Active Mbr",               "会員金額/会員人数"),
        _kpi("会員連帯率",      "Member Units Per Transaction",      "Mbr UPT",
             "Mbr Qty ÷ Active Mbr",                 "会員点数/会員TXN数"),
        _kpi("会員購入頻度",    "Purchase Frequency",                "Frequency",
             "Mbr TXNs ÷ Active Mbr",                "会員TXN数/会員人数（期間中の来店回数）"),
    ]
    return pd.DataFrame(rows, columns=_GLO_COLS)


def _style_glossary(ws):
    h_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    s_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    s_fill  = PatternFill("solid", fgColor=HEADER_BG)
    k_alt   = PatternFill("solid", fgColor="EBF3FB")
    n_font  = Font(name="Arial", size=10)

    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

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
        if is_section:
            ws.merge_cells(f"A{row}:E{row}")

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 36
    ws.row_dimensions[1].height     = 20


# ══════════════════════════════════════════════════════════════
#  スタイル & 出力
# ══════════════════════════════════════════════════════════════

def style_sheet(ws, df):
    h_font  = Font(bold=True, color=HEADER_FG, size=10, name="Arial")
    h_fill  = PatternFill("solid", fgColor=HEADER_BG)
    h_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    n_font  = Font(name="Arial", size=10)
    for cell in ws[1]:
        cell.font = h_font; cell.fill = h_fill; cell.alignment = h_align

    # 数値フォーマット（英語列名で判定）
    for i, col in enumerate(df.columns, 1):
        is_mom = "_MoM%" in col or col == "MoM%"
        is_pct = not is_mom and "%" in col
        is_flt = _is_float_col(col)
        if is_mom:
            fmt = "+0.00%;-0.00%;0.00%"
        elif is_pct:
            fmt = "0.00%"
        elif is_flt:
            fmt = "#,##0.0"
        else:
            fmt = "#,##0"
        letter = get_column_letter(i)
        for row in range(2, ws.max_row + 1):
            cell = ws[f"{letter}{row}"]
            cell.font = n_font
            if cell.value is not None and isinstance(cell.value, (int, float)):
                cell.number_format = fmt

    # 列幅
    for col_idx, col_cells in enumerate(
        ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1
    ):
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"


_OVW_FLOAT_LABELS = {"UPT", "ATV", "TF UPT", "TF ATV", "Mbr ATV", "Mbr UPT", "Frequency"}
_OVW_PCT_LABELS   = {"Year Progress%", "Year Target Reach%", "GAP%", "TF %",
                     "Month Target Reach%"}
_OVW_INT_LABELS   = {"Year YTD", "Year Target", "Month Target"}

def _style_overview(ws):
    """Overview シート専用: KPI ラベルを見てセル単位でフォーマット + Arial 適用。"""
    n_font = Font(name="Arial", size=10)
    for row in range(2, ws.max_row + 1):
        label = str(ws.cell(row=row, column=1).value or "")
        cell  = ws.cell(row=row, column=2)
        cell.font = n_font
        ws.cell(row=row, column=1).font = n_font
        if cell.value is None or not isinstance(cell.value, (int, float)):
            continue
        if label in _OVW_PCT_LABELS or label.endswith("%"):
            cell.number_format = "0.00%"
        elif "[¥100M]" in label:
            cell.number_format = "#,##0.0"
        elif label in _OVW_FLOAT_LABELS:
            cell.number_format = "#,##0.0"
        else:
            cell.number_format = "#,##0"


def write_excel(output_path, sheets: dict):
    print(f"\n  書込中: {output_path}")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for name, df in sheets.items():
            out_df = df if name == "Glossary" else _rename_df(df)
            out_df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            if name == "Glossary":
                _style_glossary(ws)
            else:
                style_sheet(ws, out_df)
                if name == "Overview":
                    _style_overview(ws)
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
    parser.add_argument("--ip-file",     default=None, help="SKU→IP マスタ")
    parser.add_argument("--budget-file", default=None, help="売上予算管理表 xlsx（省略時は自動検出）")
    args = parser.parse_args()

    print("=== YTD レポート生成ツール ===")
    paths = collect_files(args.files, args.dir, args.output)
    print(f"対象: {len(paths)} ファイル")

    df_all       = load_all(paths)
    df_ip_master = load_ip_master(args.ip_file)
    budget_data  = load_budget(args.budget_file)

    valid_months = sorted(m for m in df_all["__ym__"].unique() if m != "日付不明")
    print(f"  集計期間: {valid_months[0]} ～ {valid_months[-1]}  ({len(valid_months)} ヶ月)")

    print("  集計中...")
    sheets = {
        "Glossary":     build_glossary(),
        "Overview":     build_overview(df_all, budget_data),
        "Monthly":      build_monthly(df_all, budget_data),
        "By Store":     build_by_store(df_all, budget_data),
        "Store Monthly": build_store_monthly(df_all),
        "By SKU":       build_by_product(df_all),
    }
    if df_ip_master is not None:
        sheets = {
            "Glossary":      sheets["Glossary"],
            "Overview":      sheets["Overview"],
            "Monthly":       sheets["Monthly"],
            "By Store":      sheets["By Store"],
            "Store Monthly": sheets["Store Monthly"],
            "By IP":         build_by_ip(df_all, df_ip_master),
            "By SKU":        sheets["By SKU"],
        }
    else:
        print("  ⚠ IP マスタなし → By IP シートをスキップ")

    write_excel(args.output, sheets)


if __name__ == "__main__":
    main()
