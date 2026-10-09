import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import sys
import os

# ===================== 配置区（只改这里） =====================
# 市场选择：US / A / HK（改这一个变量即可切换市场，或通过环境变量 MARKET 传入）
MARKET = os.environ.get("MARKET", "US")

# 基础路径（自动根据 MARKET 生成输入/输出目录）
BASE_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market"
INPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo")           # 输入目录
OUTPUT_BASE_DIR = os.path.join(BASE_DIR, MARKET, "demo", "计算区间")  # 输出目录
# ============================================================


def process(INPUT_PATH, OUTPUT_PATH, INPUT_FILE1):
    in_path = Path(INPUT_PATH)
    out_path = Path(OUTPUT_PATH)
    input_file = INPUT_FILE1
    input_full = in_path / input_file

    # 输出文件名：原文件名+计算区间.xlsx
    stem = Path(input_file).stem
    output_file_name = f"{stem}计算区间.xlsx"
    
    # 如果输出目录不存在则创建
    out_path.mkdir(parents=True, exist_ok=True)
    output_full = out_path / output_file_name

    # 核心列名配置（请确保和您的Excel列名完全一致）
    COL_LATEST_PRICE = "最新价格"
    COL_LOWEST_PRICE = "历史最低价格"
    COL_HIGHEST_PRICE = "历史最高价格"
    COL_LATEST_RANGE = "最新价格区间"
    COL_LOWEST_RANGE = "历史最低区间"
    # 区间边界：2的1次方到2的30次方（对应2到1073741824）
    POWER_LIST = [2 ** n for n in range(1, 31)]  # 自动生成2、4、8...1073741824

    def calculate_price_range(price: float, highest_price: float) -> str:
        """
        计算价格所在的区间
        :param price: 目标价格（最新价格/历史最低价格）
        :param highest_price: 该行的历史最高价格
        :return: 区间字符串，如"2-4"、"≤2"、">1073741824"
        """
        # 处理空值/0值
        if pd.isna(price) or pd.isna(highest_price) or highest_price == 0:
            return "无有效数据"

        # 计算每个档位的价格阈值 H / 2, H/4 ... H/2^30
        threshold_list = [highest_price / power for power in POWER_LIST]

        # 边界情况1：价格 >= H/2 → 返回 "1-2"
        if price >= threshold_list[0]:
            return "1-2"
        # 边界情况2：价格 < H/(2^30)
        if price < threshold_list[-1]:
            return ">1073741824"

        # 找到价格所在的区间
        for i in range(len(threshold_list) - 1):
            if threshold_list[i + 1] <= price < threshold_list[i]:
                return f"{POWER_LIST[i]}-{POWER_LIST[i + 1]}"

        # 兜底返回
        return "区间计算异常"

    def beautify_excel(file_path: Path):
        """
        美化Excel格式：表头高亮、斑马纹行、自动列宽、冻结窗格、对齐优化
        :param file_path: Excel文件路径
        """
        wb = load_workbook(file_path)
        ws = wb.active

        # 样式定义
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9")
        )
        zebra_fill_even = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        zebra_fill_odd = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        max_col = ws.max_column
        max_row = ws.max_row

        # 1.表头
        for col in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment
            cell.border = thin_border

        # 2.数据行斑马纹、边框对齐
        for row in range(2, max_row + 1):
            fill = zebra_fill_even if row % 2 == 0 else zebra_fill_odd
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.fill = fill
                if col <= 3:
                    cell.alignment = left_alignment
                else:
                    cell.alignment = center_alignment

        # 3.自动列宽
        for col in range(1, max_col + 1):
            col_letter = get_column_letter(col)
            max_width = 0
            for row in range(1, max_row + 1):
                cell_val = ws.cell(row=row, column=col).value
                if cell_val is not None:
                    current_width = len(str(cell_val)) * 1.2
                    if current_width > max_width:
                        max_width = current_width
            ws.column_dimensions[col_letter].width = max_width if max_width > 8 else 10

        # 冻结首行
        ws.freeze_panes = "A2"
        wb.save(file_path)
        print(f"✅ Excel格式美化完成，已保存至: {file_path}")

    def main():
        print("=" * 50)
        print(f"开始处理文件: {input_full.name}")
        print("=" * 50)
        # 读取Excel
        try:
            df = pd.read_excel(input_full, engine="openpyxl")
            print(f"成功读取，共 {df.shape[0]} 行，{df.shape[1]} 列")
        except Exception as e:
            print(f"❌ 读取文件失败: {e}")
            return

        # 填充2^n列
        print("开始填充2^n列数据...")
        for power in POWER_LIST:
            col_name = str(power)
            if col_name not in df.columns:
                df[col_name] = pd.NA
            df[col_name] = df.apply(
                lambda row: row[COL_HIGHEST_PRICE] / power
                if pd.notna(row[COL_HIGHEST_PRICE]) and row[COL_HIGHEST_PRICE] != 0
                else pd.NA,
                axis=1
            )
        print("2^n列数据填充完成")

        # 计算区间
        print("开始计算价格区间...")
        df[COL_LATEST_RANGE] = df.apply(
            lambda row: calculate_price_range(row[COL_LATEST_PRICE], row[COL_HIGHEST_PRICE]),
            axis=1
        )
        df[COL_LOWEST_RANGE] = df.apply(
            lambda row: calculate_price_range(row[COL_LOWEST_PRICE], row[COL_HIGHEST_PRICE]),
            axis=1
        )
        print("价格区间计算完成")

        # 保存文件
        print("正在保存处理结果...")
        with pd.ExcelWriter(output_full, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="价格区间处理结果")

        beautify_excel(output_full)
        print(f"🎉 文件处理完成！输出：{output_full}\n")

    # 关键：调用内部main函数，否则逻辑完全不会跑
    main()


def get_file_list(folder):
    p = Path(folder)
    file_names = [file.name for file in p.iterdir() if file.is_file() and file.suffix.lower() in (".xlsx", ".xls")]
    return file_names


if __name__ == '__main__':
    dir_path = INPUT_DIR
    output_path = OUTPUT_BASE_DIR
    print(f"📂 输入目录: {dir_path}")
    print(f"📂 输出目录: {output_path}")
    res = get_file_list(dir_path)
    print(f"检测到待处理Excel文件共 {len(res)} 个：{res}\n")
    for file in res:
        try:
            process(dir_path, output_path, file)
        except Exception as err:    
            print(f"⚠️ 文件 {file} 处理异常跳过：{err}\n")
