"""
merge_workbooks.py
------------------
将多个 Excel 工作簿合并到一个 Sheet 中。

用法：
    python merge_workbooks.py                          # 自动合并当前目录所有 .xlsx
    python merge_workbooks.py 1月.xlsx 2月.xlsx 3月.xlsx
    python merge_workbooks.py --dir ./data --output Q1汇总.xlsx --sheet Sheet1
"""

import argparse
import glob
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ── 配置（按需修改）────────────────────────────────────────
DEFAULT_OUTPUT = "合并汇总.xlsx"
DEFAULT_SHEET_IN = 0        # 读取每个源文件的第几个 Sheet（0 = 第一个）
DEFAULT_SHEET_OUT = "汇总"  # 输出 Sheet 名称
HEADER_BG = "1F4E79"
HEADER_FG = "FFFFFF"
# ─────────────────────────────────────────────────────────


def collect_files(files: list[str], directory: str | None) -> list[Path]:
    """收集要合并的文件列表，并排序。"""
    if files:
        paths = [Path(f) for f in files]
    elif directory:
        paths = sorted(Path(directory).glob("*.xlsx"))
    else:
        paths = sorted(Path(".").glob("*.xlsx"))

    # 排除输出文件自身（如果已存在）
    paths = [p for p in paths if p.resolve() != Path(DEFAULT_OUTPUT).resolve()]

    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        sys.exit(f"[错误] 找不到文件：{', '.join(missing)}")
    if not paths:
        sys.exit("[错误] 没有找到任何 .xlsx 文件。")
    return paths


def read_sheet(path: Path, sheet_index: int) -> pd.DataFrame:
    """读取单个工作簿的指定 Sheet，返回 DataFrame。"""
    print(f"  読込中：{path.name}")
    df = pd.read_excel(path, sheet_name=sheet_index, keep_default_na=False)
    df.insert(0, "__来源文件__", path.stem)  # 可选：记录来源
    return df


def merge_dataframes(paths: list[Path], sheet_index: int) -> pd.DataFrame:
    """读取所有文件并纵向拼接，表头以第一个文件为准。"""
    dfs = []
    reference_cols = None

    for p in paths:
        df = read_sheet(p, sheet_index)
        if reference_cols is None:
            reference_cols = df.columns.tolist()
            dfs.append(df)
        else:
            # 列名对齐：缺少的列填 NaN，多余的列丢弃
            df = df.reindex(columns=reference_cols)
            dfs.append(df)

    return pd.concat(dfs, ignore_index=True)


EXCEL_MAX_ROWS = 1_048_575  # Excel 每 Sheet 最大行数（不含表头）


def apply_sheet_style(ws):
    """为 Sheet 表头添加样式并冻结首行。"""
    header_font = Font(bold=True, color=HEADER_FG, size=10)
    header_fill = PatternFill("solid", fgColor=HEADER_BG)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # 自动列宽（采样前 200 行）
    for col_idx, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=min(200, ws.max_row)), 1):
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"


def write_output(df: pd.DataFrame, output_path: str, sheet_name: str):
    """将合并后的 DataFrame 写入 Excel。
    若总行数超过 Excel 上限（1,048,576），自动按来源文件分 Sheet 写入。
    """
    total_rows = len(df)
    print(f"\n  合計：{total_rows:,} 行 × {len(df.columns)} 列")

    # ── 判断是否超过 Excel 行数上限 ──────────────────────────
    if total_rows > EXCEL_MAX_ROWS:
        print(f"  ⚠  行数が Excel の上限（{EXCEL_MAX_ROWS:,}）を超えています。")

        if "__来源文件__" in df.columns:
            # 按来源文件分 Sheet
            groups = {name: grp for name, grp in df.groupby("__来源文件__", sort=False)}
            print(f"  → 来源ファイルごとに {len(groups)} シートに分けて保存します。")
        else:
            # 按固定行数分块
            chunks = [df.iloc[i:i + EXCEL_MAX_ROWS] for i in range(0, total_rows, EXCEL_MAX_ROWS)]
            groups = {f"{sheet_name}_{idx+1}": chunk for idx, chunk in enumerate(chunks)}
            print(f"  → {len(groups)} シートに分割して保存します。")

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sname, chunk in groups.items():
                safe_name = str(sname)[:31]  # Sheet 名最大 31 文字
                chunk.to_excel(writer, sheet_name=safe_name, index=False)
                apply_sheet_style(writer.sheets[safe_name])
                print(f"    Sheet '{safe_name}'：{len(chunk):,} 行")
    else:
        # ── 正常：单 Sheet 写入 ──────────────────────────────
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            apply_sheet_style(writer.sheets[sheet_name])

    print(f"  完成！出力ファイル：{output_path}")


def main():
    parser = argparse.ArgumentParser(description="合并多个 Excel 工作簿到一个 Sheet")
    parser.add_argument("files", nargs="*", help="要合并的 Excel 文件（不填则自动扫描当前目录）")
    parser.add_argument("--dir", default=None, help="扫描指定目录下所有 .xlsx")
    parser.add_argument("--output", default=DEFAULT_OUTPUT, help=f"输出文件名（默认：{DEFAULT_OUTPUT}）")
    parser.add_argument("--sheet-in", type=int, default=DEFAULT_SHEET_IN, help="读取源文件的第几个 Sheet（0起）")
    parser.add_argument("--sheet-out", default=DEFAULT_SHEET_OUT, help=f"输出 Sheet 名（默认：{DEFAULT_SHEET_OUT}）")
    parser.add_argument("--no-source-col", action="store_true", help="不添加'来源文件'列")
    args = parser.parse_args()

    print("=== 工作簿合并工具 ===")
    paths = collect_files(args.files, args.dir)
    print(f"共找到 {len(paths)} 个文件：")
    for p in paths:
        print(f"  · {p.name}")

    df = merge_dataframes(paths, args.sheet_in)

    if args.no_source_col and "__来源文件__" in df.columns:
        df.drop(columns=["__来源文件__"], inplace=True)

    write_output(df, args.output, args.sheet_out)


if __name__ == "__main__":
    main()
