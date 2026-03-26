import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# ── 色定義 ──────────────────────────────────────────
BLUE_HEADER  = "1F4E79"
LIGHT_BLUE   = "BDD7EE"
YELLOW_INPUT = "FFF2CC"
GREEN_HEADER = "375623"
LIGHT_GREEN  = "E2EFDA"
ORANGE       = "F4B942"
GRAY_HEADER  = "595959"
LIGHT_GRAY   = "EDEDED"
WHITE        = "FFFFFF"

def style_header(cell, bg=BLUE_HEADER, font_color="FFFFFF", bold=True, size=10):
    cell.font = Font(bold=bold, color=font_color, size=size)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def style_input(cell):
    cell.fill = PatternFill("solid", fgColor=YELLOW_INPUT)
    cell.alignment = Alignment(horizontal="right", vertical="center")

def style_calc(cell):
    cell.fill = PatternFill("solid", fgColor=LIGHT_GREEN)
    cell.alignment = Alignment(horizontal="right", vertical="center")

def style_section(cell, text, bg=GRAY_HEADER):
    cell.value = text
    cell.font = Font(bold=True, color="FFFFFF", size=10)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="left", vertical="center")

def thin_border():
    s = Side(border_style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)

def apply_border(ws, min_row, max_row, min_col, max_col):
    for r in range(min_row, max_row+1):
        for c in range(min_col, max_col+1):
            ws.cell(r, c).border = thin_border()

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
DAYS   = [31,28,31,30,31,30,31,31,30,31,30,31]

# ════════════════════════════════════════════════════
# Sheet 1: 入力データ (Input)
# ════════════════════════════════════════════════════
ws1 = wb.active
ws1.title = "①入力データ"
ws1.sheet_view.showGridLines = False
ws1.column_dimensions["A"].width = 14
for col in ["B","C","D","E"]:
    ws1.column_dimensions[col].width = 14

# ── タイトル ──
ws1.merge_cells("A1:E1")
c = ws1["A1"]
c.value = "在庫金額・DOS予測モデル  ―  入力データシート"
c.font = Font(bold=True, size=13, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
c.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[1].height = 28

# 凡例
ws1.merge_cells("A2:E2")
legend = ws1["A2"]
legend.value = "🟡 黄色セル = 手入力  |  🟢 緑色セル = 自動計算（触らない）  |  ※1月・2月は実績値を入力"
legend.font = Font(size=9, italic=True)
legend.alignment = Alignment(horizontal="left")
ws1.row_dimensions[2].height = 16

# ── セクション：販売データ ──
ROW = 4
ws1.merge_cells(f"A{ROW}:E{ROW}")
style_section(ws1[f"A{ROW}"], "【販売データ】　単位：円", GRAY_HEADER)
ws1.row_dimensions[ROW].height = 20

ROW += 1
headers = ["月", "CH1 店舗", "CH2 ROBO", "CH3 B2B", "合計"]
for i, h in enumerate(headers, 1):
    c = ws1.cell(ROW, i, h)
    style_header(c)
ws1.row_dimensions[ROW].height = 18

SALES_START = ROW + 1
for mi, month in enumerate(MONTHS):
    r = SALES_START + mi
    ws1.cell(r, 1, month).alignment = Alignment(horizontal="center")
    for col in [2,3,4]:
        c = ws1.cell(r, col, 0)
        style_input(c)
        c.number_format = '#,##0'
    # 合計
    c_tot = ws1.cell(r, 5, f"=B{r}+C{r}+D{r}")
    style_calc(c_tot)
    c_tot.number_format = '#,##0'
apply_border(ws1, SALES_START, SALES_START+11, 1, 5)

# 実績/目標ラベル
for mi in range(12):
    r = SALES_START + mi
    lbl = "実績" if mi < 2 else "目標"
    ws1.cell(r, 1).value = f"{MONTHS[mi]}（{lbl}）"

# ── セクション：仕入計画 ──
ROW = SALES_START + 13
ws1.merge_cells(f"A{ROW}:E{ROW}")
style_section(ws1[f"A{ROW}"], "【仕入計画】　単位：円", GRAY_HEADER)
ws1.row_dimensions[ROW].height = 20

ROW += 1
for i, h in enumerate(headers, 1):
    c = ws1.cell(ROW, i, h)
    style_header(c)
ws1.row_dimensions[ROW].height = 18

PURCH_START = ROW + 1
for mi, month in enumerate(MONTHS):
    r = PURCH_START + mi
    ws1.cell(r, 1, MONTHS[mi]).alignment = Alignment(horizontal="center")
    for col in [2,3,4]:
        c = ws1.cell(r, col, 0)
        style_input(c)
        c.number_format = '#,##0'
    c_tot = ws1.cell(r, 5, f"=B{r}+C{r}+D{r}")
    style_calc(c_tot)
    c_tot.number_format = '#,##0'
apply_border(ws1, PURCH_START, PURCH_START+11, 1, 5)

# ── セクション：店舗数・ROBO台数 ──
ROW = PURCH_START + 13
ws1.merge_cells(f"A{ROW}:E{ROW}")
style_section(ws1[f"A{ROW}"], "【店舗数・ROBO台数】", GRAY_HEADER)
ws1.row_dimensions[ROW].height = 20

ROW += 1
store_headers = ["月", "店舗\n月初数", "新規\nOPEN", "閉店\nCLOSE", "店舗\n月末数"]
for i, h in enumerate(store_headers, 1):
    c = ws1.cell(ROW, i, h)
    style_header(c)
ws1.row_dimensions[ROW].height = 30

STORE_START = ROW + 1
for mi in range(12):
    r = STORE_START + mi
    ws1.cell(r, 1, MONTHS[mi]).alignment = Alignment(horizontal="center")
    for col in [2,3,4]:
        c = ws1.cell(r, col, 0)
        style_input(c)
        c.alignment = Alignment(horizontal="right")
    # 月末数 = 月初 + 新規 - 閉店
    c_end = ws1.cell(r, 5, f"=B{r}+C{r}-D{r}")
    style_calc(c_end)
apply_border(ws1, STORE_START, STORE_START+11, 1, 5)

# ROBO 台数（F-I 列）
ws1.column_dimensions["F"].width = 2   # spacer
ws1.column_dimensions["G"].width = 14
ws1.column_dimensions["H"].width = 14
ws1.column_dimensions["I"].width = 14

ROW_ROBO_HDR = STORE_START - 1
robo_headers = ["月", "ROBO\n月初台数", "新規\n追加台数", "ROBO\n月末台数"]
for i, h in enumerate(robo_headers, 1):
    c = ws1.cell(ROW_ROBO_HDR, 6+i, h)
    style_header(c)
ws1.row_dimensions[ROW_ROBO_HDR].height = 30

for mi in range(12):
    r = STORE_START + mi
    ws1.cell(r, 7, MONTHS[mi]).alignment = Alignment(horizontal="center")
    for col in [8,9]:
        c = ws1.cell(r, col, 0)
        style_input(c)
        c.alignment = Alignment(horizontal="right")
    c_end = ws1.cell(r, 10, f"=H{r}+I{r}")
    style_calc(c_end)
apply_border(ws1, STORE_START, STORE_START+11, 7, 10)

# ── セクション：パラメータ ──
ROW = STORE_START + 13
ws1.merge_cells(f"A{ROW}:E{ROW}")
style_section(ws1[f"A{ROW}"], "【パラメータ】　1月末データをもとに自動計算（変更可）", GRAY_HEADER)
ws1.row_dimensions[ROW].height = 20

PARAM_ROW = ROW + 1
labels_params = [
    ("CH1 期末在庫(1月実績)",    0, "円"),
    ("CH1 期末在庫(2月実績)",    0, "円"),
    ("CH2 期末在庫(1月実績)",    0, "円"),
    ("CH2 期末在庫(2月実績)",    0, "円"),
    ("CH3 期末在庫(1月実績)",    0, "円"),
    ("CH3 期末在庫(2月実績)",    0, "円"),
    ("1店舗あたり標準在庫金額",  0, "円（新規OPENに適用）"),
    ("1台ROBOあたり標準在庫金額",0, "円（新規追加に適用）"),
]
for i, (lbl, val, unit) in enumerate(labels_params):
    r = PARAM_ROW + i
    ws1.cell(r, 1, lbl).font = Font(bold=True, size=9)
    c_val = ws1.cell(r, 2, val)
    style_input(c_val)
    c_val.number_format = '#,##0'
    ws1.cell(r, 3, unit).font = Font(size=9, color="888888")
ws1.column_dimensions["B"].width = 18
apply_border(ws1, PARAM_ROW, PARAM_ROW+7, 1, 3)

# セルに名前付き範囲（Named Ranges）を使いやすくする
# パラメータセル番号を変数として保持
PARAM_CELLS = {
    "CH1_inv_jan": f"B{PARAM_ROW}",
    "CH1_inv_feb": f"B{PARAM_ROW+1}",
    "CH2_inv_jan": f"B{PARAM_ROW+2}",
    "CH2_inv_feb": f"B{PARAM_ROW+3}",
    "CH3_inv_jan": f"B{PARAM_ROW+4}",
    "CH3_inv_feb": f"B{PARAM_ROW+5}",
    "store_std":   f"B{PARAM_ROW+6}",
    "robo_std":    f"B{PARAM_ROW+7}",
}

# ── サマリ行番号をメモ ──
META = {
    "sales_start":  SALES_START,
    "purch_start":  PURCH_START,
    "store_start":  STORE_START,
    "param_cells":  PARAM_CELLS,
}

# ════════════════════════════════════════════════════
# Sheet 2: CH1 店舗
# ════════════════════════════════════════════════════
ws2 = wb.create_sheet("②CH1_店舗")
ws2.sheet_view.showGridLines = False

COL_LABELS = ["月","当月日数","月初在庫","仕入","新規OPEN調整",
              "閉店調整","当月売上","月末在庫","翌月日販","DOS（日）"]
col_widths  = [10,8,14,14,14,14,14,14,12,10]

def build_channel_sheet(ws, ch_name, bg_color):
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = f"在庫・DOS予測  ―  {ch_name}"
    c.font = Font(bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=bg_color)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:J2")
    note = ws["A2"]
    note.value = "※ 月初在庫：1月は入力データシートの実績値を参照　2月以降は前月末在庫を自動引継ぎ"
    note.font = Font(size=8, italic=True)
    note.alignment = Alignment(horizontal="left")

    for i, (lbl, w) in enumerate(zip(COL_LABELS, col_widths), 1):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = w
        c = ws.cell(3, i, lbl)
        style_header(c, bg=bg_color)
    ws.row_dimensions[3].height = 20

build_channel_sheet(ws2, "CH1 店舗", "1F4E79")

INP = "①入力データ"  # 参照シート名
SR  = META["sales_start"]
PR  = META["purch_start"]
ST  = META["store_start"]
PC  = META["param_cells"]

DATA_ROW_CH1 = 4
for mi in range(12):
    r = DATA_ROW_CH1 + mi
    sr = SR + mi   # 販売データ行
    pr = PR + mi   # 仕入データ行
    str_r = ST + mi  # 店舗数行
    ws2.row_dimensions[r].height = 18

    # A: 月
    c = ws2.cell(r, 1, MONTHS[mi])
    c.alignment = Alignment(horizontal="center")

    # B: 当月日数
    c = ws2.cell(r, 2, DAYS[mi])
    c.alignment = Alignment(horizontal="center")

    # C: 月初在庫
    if mi == 0:
        # 1月 → 入力データのパラメータから
        formula = f"='{INP}'!{PC['CH1_inv_jan']}"
    elif mi == 1:
        # 2月 → 入力データのパラメータから（2月実績）
        formula = f"='{INP}'!{PC['CH1_inv_feb']}"
    else:
        formula = f"=H{r-1}"  # 前月末在庫
    c = ws2.cell(r, 3, formula)
    c.number_format = '#,##0'
    if mi < 2:
        style_input(c)
        c.value = formula
    else:
        style_calc(c)

    # D: 仕入
    c = ws2.cell(r, 4, f"='{INP}'!B{pr}")
    style_calc(c)
    c.number_format = '#,##0'

    # E: 新規OPEN調整  = 新規OPEN数 × 1店舗標準在庫
    c = ws2.cell(r, 5, f"='{INP}'!C{str_r}*'{INP}'!{PC['store_std']}")
    style_calc(c)
    c.number_format = '#,##0'

    # F: 閉店調整（マイナス）= 閉店数 × 月初単店平均在庫
    # 月初単店平均 = C(r) / 月初店舗数
    # 月初店舗数 = 店舗月末数(前月) or 直接入力(1月は月初数)
    # 簡易：月初在庫 / 月初店舗数
    month_start_store = f"'{INP}'!B{str_r}"  # 月初店舗数
    avg_store_inv = f"IF({month_start_store}>0, C{r}/{month_start_store}, 0)"
    c = ws2.cell(r, 6, f"=-'{INP}'!D{str_r}*({avg_store_inv})")
    style_calc(c)
    c.number_format = '#,##0'

    # G: 当月売上
    c = ws2.cell(r, 7, f"='{INP}'!B{sr}")
    style_calc(c)
    c.number_format = '#,##0'

    # H: 月末在庫 = 月初 + 仕入 + 新規OPEN - 閉店 - 売上
    c = ws2.cell(r, 8, f"=C{r}+D{r}+E{r}+F{r}-G{r}")
    style_calc(c)
    c.number_format = '#,##0'
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    # I: 翌月日販
    if mi < 11:
        next_sr = SR + mi + 1
        next_days = DAYS[mi+1]
        c = ws2.cell(r, 9, f"='{INP}'!B{next_sr}/{next_days}")
    else:
        # 12月：自身の日販
        c = ws2.cell(r, 9, f"='{INP}'!B{SR+11}/{DAYS[11]}")
    style_calc(c)
    c.number_format = '#,##0'

    # J: DOS = 月末在庫 / 翌月日販
    c = ws2.cell(r, 10, f"=IF(I{r}>0, H{r}/I{r}, 0)")
    c.number_format = '0.0'
    c.fill = PatternFill("solid", fgColor="FCE4D6")
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

apply_border(ws2, DATA_ROW_CH1, DATA_ROW_CH1+11, 1, 10)

# ════════════════════════════════════════════════════
# Sheet 3: CH2 ROBO
# ════════════════════════════════════════════════════
ws3 = wb.create_sheet("③CH2_ROBO")
ws3.sheet_view.showGridLines = False

COL_LABELS3 = ["月","当月日数","月初在庫","仕入","新規ROBO調整",
               "（予備）","当月売上","月末在庫","翌月日販","DOS（日）"]
build_channel_sheet(ws3, "CH2 ROBO", "375623")
for i, (lbl, w) in enumerate(zip(COL_LABELS3, col_widths), 1):
    ws3.column_dimensions[get_column_letter(i)].width = w
    c = ws3.cell(3, i, lbl)
    style_header(c, bg="375623")

DATA_ROW_CH2 = 4
for mi in range(12):
    r = DATA_ROW_CH2 + mi
    sr = SR + mi
    pr = PR + mi
    str_r = ST + mi
    ws3.row_dimensions[r].height = 18

    ws3.cell(r, 1, MONTHS[mi]).alignment = Alignment(horizontal="center")
    ws3.cell(r, 2, DAYS[mi]).alignment = Alignment(horizontal="center")

    if mi == 0:
        formula = f"='{INP}'!{PC['CH2_inv_jan']}"
    elif mi == 1:
        formula = f"='{INP}'!{PC['CH2_inv_feb']}"
    else:
        formula = f"=H{r-1}"
    c = ws3.cell(r, 3, formula)
    c.number_format = '#,##0'
    if mi < 2:
        style_input(c); c.value = formula
    else:
        style_calc(c)

    c = ws3.cell(r, 4, f"='{INP}'!C{pr}")
    style_calc(c); c.number_format = '#,##0'

    # 新規ROBO調整 = 新追加台数 × 1台標準在庫
    c = ws3.cell(r, 5, f"='{INP}'!I{str_r}*'{INP}'!{PC['robo_std']}")
    style_calc(c); c.number_format = '#,##0'

    ws3.cell(r, 6, "―").alignment = Alignment(horizontal="center")

    c = ws3.cell(r, 7, f"='{INP}'!C{sr}")
    style_calc(c); c.number_format = '#,##0'

    c = ws3.cell(r, 8, f"=C{r}+D{r}+E{r}-G{r}")
    style_calc(c); c.number_format = '#,##0'
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor=LIGHT_GREEN)

    if mi < 11:
        c = ws3.cell(r, 9, f"='{INP}'!C{SR+mi+1}/{DAYS[mi+1]}")
    else:
        c = ws3.cell(r, 9, f"='{INP}'!C{SR+11}/{DAYS[11]}")
    style_calc(c); c.number_format = '#,##0'

    c = ws3.cell(r, 10, f"=IF(I{r}>0, H{r}/I{r}, 0)")
    c.number_format = '0.0'
    c.fill = PatternFill("solid", fgColor="FCE4D6")
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

apply_border(ws3, DATA_ROW_CH2, DATA_ROW_CH2+11, 1, 10)

# ════════════════════════════════════════════════════
# Sheet 4: CH3 B2B
# ════════════════════════════════════════════════════
ws4 = wb.create_sheet("④CH3_B2B")
ws4.sheet_view.showGridLines = False

COL_LABELS4 = ["月","当月日数","月初在庫","仕入","（調整なし）",
               "（調整なし）","当月売上","月末在庫","翌月日販","DOS（日）"]
build_channel_sheet(ws4, "CH3 B2B", "7B3F00")
for i, (lbl, w) in enumerate(zip(COL_LABELS4, col_widths), 1):
    ws4.column_dimensions[get_column_letter(i)].width = w
    c = ws4.cell(3, i, lbl)
    style_header(c, bg="7B3F00")

DATA_ROW_CH3 = 4
for mi in range(12):
    r = DATA_ROW_CH3 + mi
    sr = SR + mi
    pr = PR + mi
    ws4.row_dimensions[r].height = 18

    ws4.cell(r, 1, MONTHS[mi]).alignment = Alignment(horizontal="center")
    ws4.cell(r, 2, DAYS[mi]).alignment = Alignment(horizontal="center")

    if mi == 0:
        formula = f"='{INP}'!{PC['CH3_inv_jan']}"
    elif mi == 1:
        formula = f"='{INP}'!{PC['CH3_inv_feb']}"
    else:
        formula = f"=H{r-1}"
    c = ws4.cell(r, 3, formula)
    c.number_format = '#,##0'
    if mi < 2:
        style_input(c); c.value = formula
    else:
        style_calc(c)

    c = ws4.cell(r, 4, f"='{INP}'!D{pr}")
    style_calc(c); c.number_format = '#,##0'

    ws4.cell(r, 5, "―").alignment = Alignment(horizontal="center")
    ws4.cell(r, 6, "―").alignment = Alignment(horizontal="center")

    c = ws4.cell(r, 7, f"='{INP}'!D{sr}")
    style_calc(c); c.number_format = '#,##0'

    c = ws4.cell(r, 8, f"=C{r}+D{r}-G{r}")
    style_calc(c); c.number_format = '#,##0'
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor="FFF2CC")

    if mi < 11:
        c = ws4.cell(r, 9, f"='{INP}'!D{SR+mi+1}/{DAYS[mi+1]}")
    else:
        c = ws4.cell(r, 9, f"='{INP}'!D{SR+11}/{DAYS[11]}")
    style_calc(c); c.number_format = '#,##0'

    c = ws4.cell(r, 10, f"=IF(I{r}>0, H{r}/I{r}, 0)")
    c.number_format = '0.0'
    c.fill = PatternFill("solid", fgColor="FCE4D6")
    c.font = Font(bold=True)
    c.alignment = Alignment(horizontal="center")

apply_border(ws4, DATA_ROW_CH3, DATA_ROW_CH3+11, 1, 10)

# ════════════════════════════════════════════════════
# Sheet 5: サマリ（合計）
# ════════════════════════════════════════════════════
ws5 = wb.create_sheet("⑤サマリ（合計）")
ws5.sheet_view.showGridLines = False

ws5.merge_cells("A1:L1")
c = ws5["A1"]
c.value = "在庫金額・DOS 月次推移サマリ　（全チャネル合計）"
c.font = Font(bold=True, size=13, color="FFFFFF")
c.fill = PatternFill("solid", fgColor=BLUE_HEADER)
c.alignment = Alignment(horizontal="center", vertical="center")
ws5.row_dimensions[1].height = 28

# ヘッダー行
SUM_HEADERS = [
    "月",
    "CH1売上","CH2売上","CH3売上","合計売上",
    "CH1月末在庫","CH2月末在庫","CH3月末在庫","合計在庫",
    "CH1 DOS","CH2 DOS","CH3 DOS","合計 DOS"
]
sum_widths = [10,12,12,12,12,12,12,12,13,9,9,9,9]
for i, (h, w) in enumerate(zip(SUM_HEADERS, sum_widths), 1):
    ws5.column_dimensions[get_column_letter(i)].width = w
    c = ws5.cell(2, i, h)
    style_header(c)
ws5.row_dimensions[2].height = 20

CH1_SHEET = "②CH1_店舗"
CH2_SHEET = "③CH2_ROBO"
CH3_SHEET = "④CH3_B2B"

for mi in range(12):
    r = 3 + mi
    d4 = DATA_ROW_CH1 + mi
    d4c2 = DATA_ROW_CH2 + mi
    d4c3 = DATA_ROW_CH3 + mi
    sr = SR + mi

    ws5.row_dimensions[r].height = 18
    ws5.cell(r, 1, MONTHS[mi]).alignment = Alignment(horizontal="center")

    # 売上
    c = ws5.cell(r, 2, f"='{CH1_SHEET}'!G{d4}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 3, f"='{CH2_SHEET}'!G{d4c2}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 4, f"='{CH3_SHEET}'!G{d4c3}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 5, f"=B{r}+C{r}+D{r}")
    c.number_format = '#,##0'; c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    # 月末在庫
    c = ws5.cell(r, 6, f"='{CH1_SHEET}'!H{d4}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 7, f"='{CH2_SHEET}'!H{d4c2}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 8, f"='{CH3_SHEET}'!H{d4c3}")
    style_calc(c); c.number_format = '#,##0'
    c = ws5.cell(r, 9, f"=F{r}+G{r}+H{r}")
    c.number_format = '#,##0'; c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)

    # DOS
    c = ws5.cell(r, 10, f"='{CH1_SHEET}'!J{d4}")
    c.number_format = '0.0'; c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor="FCE4D6")
    c = ws5.cell(r, 11, f"='{CH2_SHEET}'!J{d4c2}")
    c.number_format = '0.0'; c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor="FCE4D6")
    c = ws5.cell(r, 12, f"='{CH3_SHEET}'!J{d4c3}")
    c.number_format = '0.0'; c.alignment = Alignment(horizontal="center")
    c.fill = PatternFill("solid", fgColor="FCE4D6")

    # 合計DOS = 合計在庫 / 翌月合計日販
    if mi < 11:
        next_sr = SR + mi + 1
        next_days = DAYS[mi+1]
        total_next_sales = f"('{INP}'!E{next_sr}/{next_days})"
    else:
        total_next_sales = f"('{INP}'!E{SR+11}/{DAYS[11]})"
    c = ws5.cell(r, 13, f"=IF({total_next_sales}>0, I{r}/{total_next_sales}, 0)")
    c.number_format = '0.0'
    c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor="FF6600")
    c.font = Font(bold=True, color="FFFFFF")
    c.alignment = Alignment(horizontal="center")

apply_border(ws5, 2, 14, 1, 13)

# ── 年間合計行 ──
r_total = 15
ws5.cell(r_total, 1, "年間合計").font = Font(bold=True)
ws5.cell(r_total, 1).fill = PatternFill("solid", fgColor=BLUE_HEADER)
ws5.cell(r_total, 1).font = Font(bold=True, color="FFFFFF")
for col in range(2, 6):
    c = ws5.cell(r_total, col, f"=SUM({get_column_letter(col)}3:{get_column_letter(col)}14)")
    c.number_format = '#,##0'; c.font = Font(bold=True)
    c.fill = PatternFill("solid", fgColor=LIGHT_BLUE)
for col in range(6, 14):
    ws5.cell(r_total, col, "―").alignment = Alignment(horizontal="center")
apply_border(ws5, r_total, r_total, 1, 13)

# 保存
out_path = "/home/user/DataForSales/在庫DOS予測モデル.xlsx"
wb.save(out_path)
print(f"保存完了: {out_path}")
