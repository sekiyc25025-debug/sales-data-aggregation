import pandas as pd
import re
from copy import copy
from openpyxl import load_workbook
import sys
from pathlib import Path

print("売上データ自動集計ツール")
print("------------------------")
print("処理を開始します。")

if getattr(sys, "frozen", False):
    # exeとして実行している場合
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    # Pythonスクリプトとして実行している場合
    BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

report_path = INPUT_DIR / "製品別売上レポート.xlsx"
output_path = OUTPUT_DIR / "製品別売上レポート.xlsx"

def convert_reiwa_date(text):
    match = re.match(r"令和(\d+)年(\d+)月(\d+)日", text)

    reiwa_year = int(match.group(1))
    month = int(match.group(2))
    day = int(match.group(3))

    year = reiwa_year + 2018

    return pd.Timestamp(year=year, month=month, day=day)

a_files = [
    INPUT_DIR / "発注書A①.xlsx",
    INPUT_DIR / "発注書A②.xlsx",
]

b_files = [
    INPUT_DIR / "発注書B①.csv",
    INPUT_DIR / "発注書B②.csv",
    INPUT_DIR / "発注書B③.csv",
]

a_dfs = []
b_dfs = []

print("発注書を読み込んでいます...")

for file in a_files:
    a_df = pd.read_excel(file)
    a_df = a_df.rename(columns={
        "発注日": "order_date",
        "商品名": "product_name",
        "数量": "quantity"
    })
    a_dfs.append(a_df)
    
for file in b_files:
    b_df = pd.read_csv(file)
    b_df = b_df.rename(columns={
        "発注日付": "order_date",
        "製品名": "product_name",
        "個数": "quantity"
    })
    b_df["order_date"] = b_df["order_date"].apply(convert_reiwa_date)

    b_dfs.append(b_df)

print("データを集計しています...")

df_all = pd.concat(a_dfs + b_dfs,ignore_index=True)

df_summary = (
    df_all
    .groupby(["order_date", "product_name"], as_index=False)["quantity"]
    .sum()
)

print("レポートを作成しています...")

wb = load_workbook(report_path)
ws = wb.active

start_row = 2
template_end_row = 21

# 通常行と最終行の書式を保存
normal_styles = []
last_styles = []

for col in range(1, 4):
    normal_cell = ws.cell(row=2, column=col)
    last_cell = ws.cell(row=21, column=col)

    normal_styles.append({
        "font": copy(normal_cell.font),
        "fill": copy(normal_cell.fill),
        "border": copy(normal_cell.border),
        "alignment": copy(normal_cell.alignment),
        "number_format": normal_cell.number_format,
    })

    last_styles.append({
        "font": copy(last_cell.font),
        "fill": copy(last_cell.fill),
        "border": copy(last_cell.border),
        "alignment": copy(last_cell.alignment),
        "number_format": last_cell.number_format,
    })

last_data_row = start_row + len(df_summary) - 1

for row_index, row in enumerate(
    df_summary.itertuples(index=False),
    start=start_row
):
    values = [
        row.order_date,
        row.product_name,
        row.quantity,
    ]

    for col_index, value in enumerate(values, start=1):
        cell = ws.cell(
            row=row_index,
            column=col_index,
            value=value
        )

        # 通常行の書式
        style = normal_styles[col_index - 1]

        cell.font = copy(style["font"])
        cell.fill = copy(style["fill"])
        cell.border = copy(style["border"])
        cell.alignment = copy(style["alignment"])
        cell.number_format = style["number_format"]

    # 日付だけ表示形式を指定
    ws.cell(
        row=row_index,
        column=1
    ).number_format = "yyyy/mm/dd"

for col_index in range(1, 4):
    cell = ws.cell(
        row=last_data_row,
        column=col_index
    )

    style = last_styles[col_index - 1]

    cell.border = copy(style["border"])

wb.save(output_path)

print()
print("処理が完了しました。")
print("製品別売上レポートを output フォルダに保存しました。")
print()

input("Enterキーを押すと終了します。")