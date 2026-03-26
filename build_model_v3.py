import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
DAYS   = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

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

def hdr(cell, text):
    cell.value = text
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal="center")

def yellow(cell):
    cell.fill = PatternFill("solid", fgColor="FFF2CC")

S1 = "①入力データ"

# ════════════════════════════════════════════
# Sheet1: 入力データ
# ════════════════════════════════════════════
ws1 = wb.active
ws1.title = S1
ws1.column_dimensions["A"].width = 24
for i in range(2, 15):
    ws1.column_dimensions[get_column_letter(i)].width = 15

ws1["A1"] = "在庫・DOS予測モデル v3 ― 入力データ"
ws1["A1"].font = Font(bold=True, size=12)
ws1["A2"] = "黄色=手入力  白色=既知データ（変更可）"
ws1["A2"].font = Font(size=9, color="888888")

# ── 基準値 (Row 4-8) ──
ws1["A4"] = "【基準値】"
ws1["A4"].font = Font(bold=True)

params = [
    ("Luna 1月末 実績在庫金額（円）", None),
    ("Luna 2月末 実績在庫金額（円）", None),
    ("ToC換算率（消費税10%）",        1.1),
    ("ToB換算率",                     0.72),
    ("店舗進货 当月着荷率",            0.3),
]
PARAM_ROW_START = 5
for i, (label, val) in enumerate(params):
    r = PARAM_ROW_START + i
    ws1.cell(r, 1, label)
    c = ws1.cell(r, 2)
    c.number_format = '#,##0' if val is None else '0.00'
    if val is not None:
        c.value = val
    else:
        yellow(c)

# Named param rows
ROW_INV1 = PARAM_ROW_START      # Luna 1月末在庫
ROW_INV2 = PARAM_ROW_START + 1  # Luna 2月末在庫
ROW_TOC  = PARAM_ROW_START + 2  # ToC換算率  B7
ROW_TOB  = PARAM_ROW_START + 3  # ToB換算率  B8
ROW_LT   = PARAM_ROW_START + 4  # 当月着荷率 B9
box(ws1, PARAM_ROW_START, 1, PARAM_ROW_START+4, 2)

# ── 月ヘッダー (Row 11) ──
HDR = 11
ws1.cell(HDR, 1, "月").font = Font(bold=True)
for mi in range(12):
    hdr(ws1.cell(HDR, 2+mi), MONTHS[mi])

# ── 販売予算 (Row 13-19) ──
ws1.cell(12, 1, "【販売予算（Sell Out、税込）】").font = Font(bold=True)

SALES_LABELS = [
    ("店舗（既存）",    STORE_SALES, False),
    ("店舗（新規Open）",None,        True),
    ("ROBO",            ROBO_SALES,  False),
    ("B2B",             None,        True),
]
SROW = {}
R = 13
for label, data, need_input in SALES_LABELS:
    ws1.cell(R, 1, label)
    SROW[label] = R
    for mi in range(12):
        c = ws1.cell(R, 2+mi)
        c.number_format = '#,##0'
        if data:
            c.value = data[mi]
        if need_input:
            yellow(c)
    R += 1

# 合計行（ToC・ToB・全体）
for key, expr in [
    ("合計ToC", lambda mi: f"=B{SROW['店舗（既存）']+mi-mi}+C{SROW['店舗（既存）']}"),  # placeholder
    ("合計ToB", None),
    ("合計全体", None),
]:
    pass  # build below

ws1.cell(R, 1, "販売合計（ToC）")
ws1.cell(R, 1).font = Font(bold=True)
SROW["合計ToC"] = R
for mi in range(12):
    col = get_column_letter(2+mi)
    f = (f"={col}{SROW['店舗（既存）']}"
         f"+{col}{SROW['店舗（新規Open）']}"
         f"+{col}{SROW['ROBO']}")
    c = ws1.cell(R, 2+mi, f)
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

ws1.cell(R, 1, "販売合計（ToB）")
ws1.cell(R, 1).font = Font(bold=True)
SROW["合計ToB"] = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws1.cell(R, 2+mi, f"={col}{SROW['B2B']}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1

ws1.cell(R, 1, "販売合計（全体）")
ws1.cell(R, 1).font = Font(bold=True)
SROW["合計全体"] = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws1.cell(R, 2+mi,
                 f"={col}{SROW['合計ToC']}+{col}{SROW['合計ToB']}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1
box(ws1, 13, 1, R-1, 13)

# ── 進货・EC移管 (Row R+1 ..) ──
R += 1
ws1.cell(R, 1, "【進货計画・EC移管】").font = Font(bold=True)
R += 1

ws1.cell(R, 1, "店舗 進货計画")
PROW = R  # store purchase row
for mi in range(12):
    c = ws1.cell(R, 2+mi, STORE_PURCH[mi])
    c.number_format = '#,##0'
R += 1

ws1.cell(R, 1, "EC移管金額（当月全額着荷）")
ECROW = R
for mi in range(12):
    c = ws1.cell(R, 2+mi)
    c.number_format = '#,##0'
    yellow(c)
    if mi == 0:
        c.value = 147846918
R += 1

# 入庫合計: 進货×当月着荷率 + 前月進货×(1-着荷率) + EC移管
ws1.cell(R, 1, "入庫合計（着荷ベース）")
ws1.cell(R, 1).font = Font(bold=True)
INROW = R  # inflow row
for mi in range(12):
    col = get_column_letter(2+mi)
    if mi == 0:
        f = (f"={col}{PROW}*$B${ROW_LT}"
             f"+{col}{ECROW}")
    else:
        prev = get_column_letter(2+mi-1)
        f = (f"={col}{PROW}*$B${ROW_LT}"
             f"+{prev}{PROW}*(1-$B${ROW_LT})"
             f"+{col}{ECROW}")
    c = ws1.cell(R, 2+mi, f)
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1
box(ws1, PROW, 1, R-1, 13)

# ════════════════════════════════════════════
# Sheet2: 予測計算
# ════════════════════════════════════════════
ws2 = wb.create_sheet("②予測計算")
ws2.column_dimensions["A"].width = 26
for i in range(2, 15):
    ws2.column_dimensions[get_column_letter(i)].width = 15

ws2["A1"] = "在庫・DOS 月次予測計算"
ws2["A1"].font = Font(bold=True, size=12)
ws2["A2"] = "1月末・2月末は実績参照。3月以降は自動計算。"
ws2["A2"].font = Font(size=9, color="888888")

for mi in range(12):
    hdr(ws2.cell(3, 2+mi), MONTHS[mi])

S = S1

R = 5
ws2.cell(R, 1, "【月次消費量（原価換算）】").font = Font(bold=True)
R += 1

# 消費量 = ToC販売/換算率 + ToB販売/換算率
# ★ 換算率は '①入力データ'!$B$ROW_TOC / ROW_TOB で参照
ws2.cell(R, 1, "ToC消費（原価換算）")
CONSUME_TOC_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    f = f"='{S}'!{col}{SROW['合計ToC']}/'{S}'!$B${ROW_TOC}"
    c = ws2.cell(R, 2+mi, f)
    c.number_format = '#,##0'
R += 1

ws2.cell(R, 1, "ToB消費（原価換算）")
CONSUME_TOB_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    f = f"='{S}'!{col}{SROW['合計ToB']}/'{S}'!$B${ROW_TOB}"
    c = ws2.cell(R, 2+mi, f)
    c.number_format = '#,##0'
R += 1

ws2.cell(R, 1, "消費合計（原価換算）")
ws2.cell(R, 1).font = Font(bold=True)
CONSUME_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi,
                 f"={col}{CONSUME_TOC_ROW}+{col}{CONSUME_TOB_ROW}")
    c.number_format = '#,##0'
    c.font = Font(bold=True)
R += 1
box(ws2, CONSUME_TOC_ROW, 1, CONSUME_ROW, 13)

R += 1
ws2.cell(R, 1, "【在庫ロールフォワード】").font = Font(bold=True)
R += 1

ws2.cell(R, 1, "入庫合計（参照）")
INFLOW_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi, f"='{S}'!{col}{INROW}")
    c.number_format = '#,##0'
R += 1

ws2.cell(R, 1, "月末在庫")
ws2.cell(R, 1).font = Font(bold=True)
INV_END_ROW = R
for mi in range(12):
    col = get_column_letter(2+mi)
    c = ws2.cell(R, 2+mi)
    c.number_format = '#,##0'
    c.font = Font(bold=True)
    if mi == 0:
        # 1月末 = 実績
        c.value = f"='{S}'!$B${ROW_INV1}"
    elif mi == 1:
        # 2月末 = 実績
        c.value = f"='{S}'!$B${ROW_INV2}"
    else:
        # 3月以降: 前月末 + 当月入庫 - 当月消費
        prev = get_column_letter(2+mi-1)
        c.value = (f"={prev}{INV_END_ROW}"
                   f"+{col}{INFLOW_ROW}"
                   f"-{col}{CONSUME_ROW}")
R += 1
box(ws2, INFLOW_ROW, 1, INV_END_ROW, 13)

R += 1
ws2.cell(R, 1, "【DOS計算（月数単位）】").font = Font(bold=True)
R += 1

# 2027年1月消費量 = 2026年1月実績と同額
NEXT_JAN_CONSUME = f"B{CONSUME_ROW}"  # = January 2026 consumption

ws2.cell(R, 1, "DOS（月数）")
ws2.cell(R, 1).font = Font(bold=True)
DOS_ROW = R

for mi in range(12):
    col  = get_column_letter(2+mi)
    inv  = f"{col}{INV_END_ROW}"

    def cm(offset):
        """月次消費量セル参照 (offset=0→当月, 1→翌月...)"""
        if mi + offset < 12:
            return f"{get_column_letter(2+mi+offset)}{CONSUME_ROW}"
        else:
            return NEXT_JAN_CONSUME  # 2027年1月

    if mi <= 8:
        # 1〜9月: 積み上げ方式（翌月・翌々月・翌々々月）
        c1, c2, c3 = cm(1), cm(2), cm(3)
        formula = (
            f"=IF({c1}=0,0,"
            f"IF({inv}<{c1},"
            f"{inv}/{c1},"
            f"IF({inv}<{c1}+{c2},"
            f"1+({inv}-{c1})/{c2},"
            f"2+({inv}-{c1}-{c2})/{c3})))"
        )
    elif mi == 9:
        # 10月: 残り=11月・12月 → 平均
        c_nov = cm(1)
        c_dec = cm(2)
        formula = (
            f"=IF(({c_nov}+{c_dec})=0,0,"
            f"{inv}/(({c_nov}+{c_dec})/2))"
        )
    elif mi == 10:
        # 11月: 残り=12月のみ
        c_dec = cm(1)
        formula = f"=IF({c_dec}=0,0,{inv}/{c_dec})"
    else:
        # 12月: 残り=2027年1月（=2026年1月消費量）
        formula = f"=IF({NEXT_JAN_CONSUME}=0,0,{inv}/{NEXT_JAN_CONSUME})"

    c = ws2.cell(R, 2+mi, formula)
    c.number_format = '0.00'
    c.font = Font(bold=True)

box(ws2, DOS_ROW, 1, DOS_ROW, 13)

# ════════════════════════════════════════════
# Sheet3: サマリ
# ════════════════════════════════════════════
ws3 = wb.create_sheet("③サマリ")
ws3.column_dimensions["A"].width = 24
for i in range(2, 15):
    ws3.column_dimensions[get_column_letter(i)].width = 14

ws3["A1"] = "サマリ ― 月次在庫・DOS推移"
ws3["A1"].font = Font(bold=True, size=12)

for mi in range(12):
    hdr(ws3.cell(3, 2+mi), MONTHS[mi])

CAL = "②予測計算"
summary = [
    ("販売合計（全体）",   f"'{S1}'",  SROW["合計全体"],  '#,##0', False),
    ("  ToC販売",          f"'{S1}'",  SROW["合計ToC"],   '#,##0', False),
    ("  ToB販売",          f"'{S1}'",  SROW["合計ToB"],   '#,##0', False),
    ("入庫合計",           f"'{S1}'",  INROW,             '#,##0', False),
    ("  店舗進货",         f"'{S1}'",  PROW,              '#,##0', False),
    ("  EC移管",           f"'{S1}'",  ECROW,             '#,##0', False),
    ("消費合計（原価換算）",f"'{CAL}'", CONSUME_ROW,       '#,##0', False),
    ("月末在庫金額",       f"'{CAL}'", INV_END_ROW,       '#,##0', True),
    ("DOS（月数）",        f"'{CAL}'", DOS_ROW,           '0.00',  True),
]

sr = 5
for label, sheet, src_row, fmt, bold in summary:
    ws3.cell(sr, 1, label)
    if bold:
        ws3.cell(sr, 1).font = Font(bold=True)
    for mi in range(12):
        col = get_column_letter(2+mi)
        c = ws3.cell(sr, 2+mi, f"={sheet}!{col}{src_row}")
        c.number_format = fmt
        if bold:
            c.font = Font(bold=True)
    sr += 1

box(ws3, 4, 1, sr-1, 13)

out = "/home/user/DataForSales/在庫DOS予測モデル_v3.xlsx"
wb.save(out)
print(f"保存完了: {out}")
