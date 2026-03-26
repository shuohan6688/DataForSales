import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
DAYS   = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

# 既知データ
STORE_SALES = [1699940255,1313219737,1470783378,1801379308,1874294789,1911750692,
               1925967529,1893433375,1711900823,1663896416,1523414395,1349930172]
ROBO_SALES  = [73558644,56824740,63642725,77948044,81103193,82723959,
               83339140,81931345,74076193,71998979,65920138,58413248]
STORE_PURCH = [2317652861,850096243,1315131747,1472696605,1460578854,2024355831,
               1835776567,1529507202,1502459212,1230694625,806516283,773997069]

def thin():
    s = Side(border_style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def box(ws, r1, c1, r2, c2):
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r, c).border = thin()

S1 = "①入力データ"

# ════════════════════════════════════════════
# Sheet1: 入力データ
# ════════════════════════════════════════════
ws1 = wb.active
ws1.title = S1
ws1.column_dimensions["A"].width = 22
for i in range(2, 15):
    ws1.column_dimensions[get_column_letter(i)].width = 16

# Row 1: タイトル
ws1["A1"] = "在庫・DOS予測モデル v3 ― 入力データ"
ws1["A1"].font = Font(bold=True, size=12)

# Row 2: 凡例
ws1["A2"] = "黄色=手入力  白色=自動計算/既知データ"
ws1["A2"].font = Font(size=9, italic=True, color="888888")

# ── 基準値 ──
ws1["A4"] = "【基準値】"
ws1["A4"].font = Font(bold=True)
ws1["A5"] = "Luna 2月末 実績在庫金額（円）"
ws1["B5"].fill = PatternFill("solid", fgColor="FFF2CC")
ws1["B5"].number_format = '#,##0'
ws1["A6"] = "ToC換算率（消費税）"
ws1["B6"] = 1.1
ws1["A7"] = "ToB換算率"
ws1["B7"] = 0.72
ws1["A8"] = "店舗進货 当月着荷率"
ws1["B8"] = 0.3
box(ws1, 5, 1, 8, 2)

# ── 月ヘッダー（共通）Row 10 ──
HDR_ROW = 10
ws1.cell(HDR_ROW, 1, "月").font = Font(bold=True)
for mi in range(12):
    c = ws1.cell(HDR_ROW, 2+mi, MONTHS[mi])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

# ── Section: 販売予算 ──
R = 12
ws1.cell(R, 1, "【販売予算（Sell Out）】").font = Font(bold=True)
R += 1

labels_sales = [
    ("店舗（既存）", STORE_SALES, False),
    ("店舗（新規Open）", None, True),
    ("ROBO", ROBO_SALES, False),
    ("B2B", None, True),
]
SALES_ROWS = {}
for label, data, need_input in labels_sales:
    ws1.cell(R, 1, label)
    SALES_ROWS[label] = R
    for mi in range(12):
        c = ws1.cell(R, 2+mi)
        if data:
            c.value = data[mi]
        elif need_input:
            c.fill = PatternFill("solid", fgColor="FFF2CC")
        c.number_format = '#,##0'
    R += 1

# 合計行
ws1.cell(R, 1, "販売合計（ToC）")
ws1.cell(R, 1).font = Font(bold=True)
SALES_ROWS["合計ToC"] = R
r_exist = SALES_ROWS["店舗（既存）"]
r_new   = SALES_ROWS["店舗（新規Open）"]
r_robo  = SALES_ROWS["ROBO"]
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws1.cell(R, 2+mi, f"={col}{r_exist}+{col}{r_new}+{col}{r_robo}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

ws1.cell(R, 1, "販売合計（ToB = B2B）")
ws1.cell(R, 1).font = Font(bold=True)
SALES_ROWS["合計ToB"] = R
r_b2b = SALES_ROWS["B2B"]
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws1.cell(R, 2+mi, f"={col}{r_b2b}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

# 販売合計（全体）
ws1.cell(R, 1, "販売合計（全体）")
ws1.cell(R, 1).font = Font(bold=True)
SALES_ROWS["合計全体"] = R
r_toc = SALES_ROWS["合計ToC"]
r_tob = SALES_ROWS["合計ToB"]
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws1.cell(R, 2+mi, f"={col}{r_toc}+{col}{r_tob}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

box(ws1, SALES_ROWS["店舗（既存）"], 1, R-1, 13)

# ── Section: 進货・EC移管 ──
R += 1
ws1.cell(R, 1, "【進货計画・EC移管】").font = Font(bold=True)
R += 1

ws1.cell(R, 1, "店舗 進货計画")
PURCH_ROW = R
for mi in range(12):
    c = ws1.cell(R, 2+mi, STORE_PURCH[mi])
    c.number_format = '#,##0'
R += 1

ws1.cell(R, 1, "EC移管金額（Luna着荷）")
EC_ROW = R
for mi in range(12):
    c = ws1.cell(R, 2+mi)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.number_format = '#,##0'
    if mi == 0:
        c.value = 147846918
R += 1

ws1.cell(R, 1, "入庫合計")
ws1.cell(R, 1).font = Font(bold=True)
TOTAL_IN_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    # 店舗進货×当月着荷率 + 前月店舗進货×(1-当月着荷率) + EC移管
    if mi == 0:
        formula = (f"={col}{PURCH_ROW}*$B$8"
                   f"+{col}{EC_ROW}")
    else:
        prev_col = get_column_letter(2+mi-1)
        formula = (f"={col}{PURCH_ROW}*$B$8"
                   f"+{prev_col}{PURCH_ROW}*(1-$B$8)"
                   f"+{col}{EC_ROW}")
    c = ws1.cell(R, 2+mi, formula)
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

box(ws1, PURCH_ROW, 1, R-1, 13)

# ── Section: 新規開店計画 ──
R += 1
ws1.cell(R, 1, "【新規開店計画】").font = Font(bold=True)
R += 1

ws1.cell(R, 1, "新規開店数（店）")
NEWSTORE_ROW = R
for mi in range(12):
    c = ws1.cell(R, 2+mi)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.number_format = '0'
R += 1

ws1.cell(R, 1, "累計店舗数（既存+新規）")
CUMSTORE_ROW = R
for mi in range(12):
    c = ws1.cell(R, 2+mi)
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.number_format = '0'
R += 1

box(ws1, NEWSTORE_ROW, 1, CUMSTORE_ROW, 13)

# ════════════════════════════════════════════
# Sheet2: 予測計算
# ════════════════════════════════════════════
ws2 = wb.create_sheet("②予測計算")
ws2.column_dimensions["A"].width = 24
for i in range(2, 15):
    ws2.column_dimensions[get_column_letter(i)].width = 15

ws2["A1"] = "在庫・DOS 月次予測計算"
ws2["A1"].font = Font(bold=True, size=12)

# 月ヘッダー
for mi in range(12):
    c = ws2.cell(3, 2+mi, MONTHS[mi])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

S = S1  # 参照シート名

R = 5
ws2.cell(R, 1, "【在庫ロールフォワード】").font = Font(bold=True)
R += 1

# 月初在庫
ws2.cell(R, 1, "月初在庫")
INV_BEGIN_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi)
    c.number_format = '#,##0'
    if mi == 0:
        # 1月月初 = 不明（前年12月末）→ 0 or skip
        c.value = 0
        c.fill = PatternFill("solid", fgColor="FFF2CC")
    elif mi == 1:
        # 2月月初 = 1月末在庫
        c.value = f"={col}{INV_BEGIN_ROW+5}"  # will be INV_END_ROW for prev month
        # placeholder, fix after INV_END defined
        pass
    else:
        prev_col = get_column_letter(2+mi-1)
        pass  # fix after
R += 1

# 入庫合計（参照）
ws2.cell(R, 1, "入庫合計")
INFLOW_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi, f"='{S}'!{col}{TOTAL_IN_ROW}")
    c.number_format = '#,##0'
R += 1

# Sell Out消費量（原価換算）
ws2.cell(R, 1, "Sell Out消費（原価換算）")
CONSUME_ROW = R
r_toc = SALES_ROWS["合計ToC"]
r_tob = SALES_ROWS["合計ToB"]
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi, f"='{S}'!{col}{r_toc}/$B$6+'{S}'!{col}{r_tob}/$B$7")
    c.number_format = '#,##0'
R += 1

# 月末在庫
ws2.cell(R, 1, "月末在庫")
ws2.cell(R, 1).font = Font(bold=True)
INV_END_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    if mi < 2:
        # 1月, 2月 = 実績 → 2月末はパラメータ参照
        if mi == 0:
            c = ws2.cell(R, 2+mi)
            c.fill = PatternFill("solid", fgColor="FFF2CC")
            c.number_format = '#,##0'
        else:
            c = ws2.cell(R, 2+mi, f"='{S}'!B5")
            c.number_format = '#,##0'
    else:
        c = ws2.cell(R, 2+mi,
                     f"={col}{INV_BEGIN_ROW}+{col}{INFLOW_ROW}-{col}{CONSUME_ROW}")
        c.number_format = '#,##0'
        c.font = Font(bold=True)
R += 1

# 月初在庫を再設定（月末在庫の行が確定したので）
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(INV_BEGIN_ROW, 2+mi)
    if mi == 0:
        c.value = 0  # 1月月初は不明 → 手入力 or 0
        c.fill = PatternFill("solid", fgColor="FFF2CC")
    elif mi <= 1:
        prev_col = get_column_letter(2+mi-1)
        c.value = f"={prev_col}{INV_END_ROW}"
    else:
        prev_col = get_column_letter(2+mi-1)
        c.value = f"={prev_col}{INV_END_ROW}"
    c.number_format = '#,##0'

box(ws2, INV_BEGIN_ROW, 1, INV_END_ROW, 13)

# ── DOS計算 ──
R += 1
ws2.cell(R, 1, "【DOS計算（月数単位）】").font = Font(bold=True)
R += 1

# 月次消費量（原価換算）参照行
ws2.cell(R, 1, "月次消費量（原価換算）")
MONTHLY_CONSUME_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi, f"={col}{CONSUME_ROW}")
    c.number_format = '#,##0'
R += 1

# 2027年1月消費量（=2026年1月実績）
ws2.cell(R, 1, "2027年1月消費量（参考）")
NEXT_JAN_ROW = R
c = ws2.cell(R, 2, f"=B{MONTHLY_CONSUME_ROW}")  # = 2026年1月と同額
c.number_format = '#,##0'
ws2["A" + str(R)].font = Font(size=9, color="888888")
R += 1

# DOS
ws2.cell(R, 1, "DOS（月数）")
ws2.cell(R, 1).font = Font(bold=True)
DOS_ROW = R

for mi in range(12):
    col = get_column_letter(2+mi)
    inv = f"{col}{INV_END_ROW}"

    if mi <= 8:
        # 1月~9月: 積み上げ方式（翌月, 翌々月, 翌々々月）
        # M+1, M+2, M+3
        m1 = get_column_letter(2+mi+1) if mi+1 < 12 else None
        m2 = get_column_letter(2+mi+2) if mi+2 < 12 else None
        m3 = get_column_letter(2+mi+3) if mi+3 < 12 else None

        if m1 and m2 and m3:
            c1 = f"{m1}{MONTHLY_CONSUME_ROW}"
            c2 = f"{m2}{MONTHLY_CONSUME_ROW}"
            c3 = f"{m3}{MONTHLY_CONSUME_ROW}"
            formula = (
                f"=IF({inv}<{c1},"
                f"{inv}/{c1},"
                f"IF({inv}<{c1}+{c2},"
                f"1+({inv}-{c1})/{c2},"
                f"2+({inv}-{c1}-{c2})/{c3}))"
            )
        elif m1 and m2:
            # m3 is None (11月参照 = 2027年1月)
            c1 = f"{m1}{MONTHLY_CONSUME_ROW}"
            c2 = f"{m2}{MONTHLY_CONSUME_ROW}"
            c3 = f"B{NEXT_JAN_ROW}"
            formula = (
                f"=IF({inv}<{c1},"
                f"{inv}/{c1},"
                f"IF({inv}<{c1}+{c2},"
                f"1+({inv}-{c1})/{c2},"
                f"2+({inv}-{c1}-{c2})/{c3}))"
            )
        else:
            # fallback
            c1 = f"{m1}{MONTHLY_CONSUME_ROW}" if m1 else f"B{NEXT_JAN_ROW}"
            formula = f"=IF({c1}>0,{inv}/{c1},0)"
    elif mi == 9:
        # 10月: 残り11月,12月 → 平均(11月,12月)
        c_nov = f"{get_column_letter(2+10)}{MONTHLY_CONSUME_ROW}"
        c_dec = f"{get_column_letter(2+11)}{MONTHLY_CONSUME_ROW}"
        avg = f"({c_nov}+{c_dec})/2"
        formula = f"=IF({avg}>0,{inv}/{avg},0)"
    elif mi == 10:
        # 11月: 残り12月のみ
        c_dec = f"{get_column_letter(2+11)}{MONTHLY_CONSUME_ROW}"
        formula = f"=IF({c_dec}>0,{inv}/{c_dec},0)"
    else:
        # 12月: 2027年1月（=2026年1月実績）
        c_jan = f"B{NEXT_JAN_ROW}"
        formula = f"=IF({c_jan}>0,{inv}/{c_jan},0)"

    c = ws2.cell(R, 2+mi, formula)
    c.number_format = '0.0'
    c.font = Font(bold=True)
R += 1

box(ws2, MONTHLY_CONSUME_ROW, 1, DOS_ROW, 13)

# ════════════════════════════════════════════
# Sheet3: サマリ
# ════════════════════════════════════════════
ws3 = wb.create_sheet("③サマリ")
ws3.column_dimensions["A"].width = 22
for i in range(2, 15):
    ws3.column_dimensions[get_column_letter(i)].width = 14

ws3["A1"] = "サマリダッシュボード ― 月次在庫・DOS推移"
ws3["A1"].font = Font(bold=True, size=12)

for mi in range(12):
    c = ws3.cell(3, 2+mi, MONTHS[mi])
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

CAL = "②予測計算"
rows_summary = [
    ("Sell Out 販売合計",     f"'{S}'!COL{SALES_ROWS['合計全体']}", '#,##0'),
    ("  うちToC",             f"'{S}'!COL{SALES_ROWS['合計ToC']}", '#,##0'),
    ("  うちToB",             f"'{S}'!COL{SALES_ROWS['合計ToB']}", '#,##0'),
    ("入庫合計",              f"'{S}'!COL{TOTAL_IN_ROW}", '#,##0'),
    ("  店舗進货",            f"'{S}'!COL{PURCH_ROW}", '#,##0'),
    ("  EC移管",              f"'{S}'!COL{EC_ROW}", '#,##0'),
    ("消費量（原価換算）",    f"'{CAL}'!COL{CONSUME_ROW}", '#,##0'),
    ("月末在庫金額",          f"'{CAL}'!COL{INV_END_ROW}", '#,##0'),
    ("DOS（月数）",           f"'{CAL}'!COL{DOS_ROW}", '0.0'),
]

sr = 5
for label, ref_template, fmt in rows_summary:
    ws3.cell(sr, 1, label)
    if label in ["月末在庫金額", "DOS（月数）"]:
        ws3.cell(sr, 1).font = Font(bold=True)
    for mi in range(12):
        col = get_column_letter(2+mi)
        ref = ref_template.replace("COL", col)
        c = ws3.cell(sr, 2+mi, f"={ref}")
        c.number_format = fmt
        if label in ["月末在庫金額", "DOS（月数）"]:
            c.font = Font(bold=True)
    sr += 1

box(ws3, 4, 1, sr-1, 13)

# 保存
out = "/home/user/DataForSales/在庫DOS予測モデル_v3.xlsx"
wb.save(out)
print(f"保存完了: {out}")
