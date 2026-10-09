import os
import glob
import re
from copy import copy
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# ===== 配置 =====
# 通过环境变量接收输入目录
INPUT_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market/US/demo/计算区间/双区间相同再算幅度/标黄/已标注"#os.environ.get("PROCESS_INPUT_DIR", "")
SHEET1_NAME = "价格区间处理结果"
SHEET2_NAME = "双区间再分表格"
OUTPUT_SUFFIX = "_已精减"

# IBKR 输出目录
IBKR_OUTPUT_DIR = r"/home/wanguangji/IBKR_deepseek/input_data"

# 黄 / 红 背景色 RGB（6位，不含alpha）
YELLOW_RGB = {"FFFF00"}
RED_RGB = {"FF0000", "FFC000"}


def normalize_rgb(rgb):
    """把各种格式的 rgb 统一成6位大写，如 'FFFFFF00' -> 'FFFF00'"""
    if rgb is None:
        return None
    rgb = str(rgb).upper()
    if len(rgb) == 8:
        rgb = rgb[2:]
    return rgb


def is_yellow_or_red(cell):
    """判断单元格背景是否为黄色或红色（solid fill）"""
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return False
    rgb = normalize_rgb(fill.start_color.rgb)
    if rgb is None:
        return False
    if rgb in YELLOW_RGB or rgb in RED_RGB:
        return True
    return False


def get_highlighted_codes(ws_sheet2):
    """从 Sheet2 的 A 列获取标黄/红背景的股票代码（保持顺序）"""
    codes = []
    for row in range(2, ws_sheet2.max_row + 1):
        cell = ws_sheet2.cell(row=row, column=1)
        if cell.value is None:
            continue
        if is_yellow_or_red(cell):
            codes.append(str(cell.value).strip())
    return codes


def adjust_formula_row(formula, row_map):
    """
    把公式里的行号按 row_map 重新映射。
    row_map: {旧行号: 新行号}
    """
    if not isinstance(formula, str) or not formula.startswith("="):
        return formula

    def replace_row(match):
        col = match.group(1)
        old_row = int(match.group(2))
        new_row = row_map.get(old_row, old_row)
        return f"{col}{new_row}"

    pattern = re.compile(r'([A-Z]{1,3})(\d+)')
    return pattern.sub(replace_row, formula)


def copy_sheet_filtered(ws_src, ws_dst, keep_codes):
    """
    把 ws_src 中 A 列在 keep_codes 中的行复制到 ws_dst。
    保留值和样式；公式按新行号重新映射。
    """
    # 1. 找出要保留的旧行号
    keep_rows = []
    for row in range(2, ws_src.max_row + 1):
        val = ws_src.cell(row=row, column=1).value
        if val is None:
            continue
        if str(val).strip() in keep_codes:
            keep_rows.append(row)

    # 2. 建立 旧行号 -> 新行号 映射（新表第1行是表头）
    row_map = {}
    for new_idx, old_row in enumerate(keep_rows, start=2):
        row_map[old_row] = new_idx

    # 3. 复制表头（第1行）
    max_col = ws_src.max_column
    for col in range(1, max_col + 1):
        src_cell = ws_src.cell(row=1, column=col)
        dst_cell = ws_dst.cell(row=1, column=col)
        dst_cell.value = src_cell.value
        if src_cell.has_style:
            dst_cell.font = copy(src_cell.font)
            dst_cell.fill = copy(src_cell.fill)
            dst_cell.border = copy(src_cell.border)
            dst_cell.alignment = copy(src_cell.alignment)
            dst_cell.number_format = src_cell.number_format

    # 4. 复制保留的数据行
    for old_row, new_row in row_map.items():
        for col in range(1, max_col + 1):
            src_cell = ws_src.cell(row=old_row, column=col)
            dst_cell = ws_dst.cell(row=new_row, column=col)

            value = src_cell.value
            # 如果是公式，按 row_map 调整行号
            if isinstance(value, str) and value.startswith("="):
                value = adjust_formula_row(value, row_map)
            dst_cell.value = value

            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.border = copy(src_cell.border)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.number_format = src_cell.number_format

    # 5. 复制列宽
    for col in range(1, max_col + 1):
        col_letter = get_column_letter(col)
        if col_letter in ws_src.column_dimensions:
            ws_dst.column_dimensions[col_letter].width = \
                ws_src.column_dimensions[col_letter].width


def clear_jingjian_files(directory):
    """删除目录中所有文件名包含'已精减'的xlsx文件"""
    if not os.path.exists(directory):
        return
    for f in glob.glob(os.path.join(directory, "*已精减*.xlsx")):
        try:
            os.remove(f)
            print(f"  已删除旧文件: {f}")
        except Exception as e:
            print(f"  删除失败 {f}: {e}")


def save_to_directory(new_wb, filename, directory):
    """保存文件到指定目录（目录不存在则创建）"""
    os.makedirs(directory, exist_ok=True)
    out_path = os.path.join(directory, filename)
    new_wb.save(out_path)
    print(f"  已保存: {out_path}")


def process_file(filepath, jingjian_dir, ibkr_dir):
    print(f"\n处理文件: {filepath}")
    wb = load_workbook(filepath)

    if SHEET1_NAME not in wb.sheetnames:
        print(f"  跳过：未找到 Sheet '{SHEET1_NAME}'")
        return
    if SHEET2_NAME not in wb.sheetnames:
        print(f"  跳过：未找到 Sheet '{SHEET2_NAME}'")
        return

    ws1_src = wb[SHEET1_NAME]
    ws2_src = wb[SHEET2_NAME]

    # 从 Sheet2 的 A 列获取标黄/红个股
    keep_codes_list = get_highlighted_codes(ws2_src)
    keep_codes = set(keep_codes_list)
    print(f"  Sheet2 中标黄/红的个股数量: {len(keep_codes_list)}")
    if not keep_codes:
        print("  没有标黄/红的个股，跳过该文件")
        return
    print(f"  个股列表: {keep_codes_list}")

    # 新建工作簿
    new_wb = Workbook()
    default_ws = new_wb.active
    new_wb.remove(default_ws)

    ws1_dst = new_wb.create_sheet(title=SHEET1_NAME)
    copy_sheet_filtered(ws1_src, ws1_dst, keep_codes)

    ws2_dst = new_wb.create_sheet(title=SHEET2_NAME)
    copy_sheet_filtered(ws2_src, ws2_dst, keep_codes)

    new_wb.calculation.fullCalcOnLoad = True

    # 生成输出文件名：原名_已精减.xlsx
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    out_filename = f"{base_name}{OUTPUT_SUFFIX}.xlsx"

    # 保存到两个目标目录
    save_to_directory(new_wb, out_filename, jingjian_dir)
    save_to_directory(new_wb, out_filename, ibkr_dir)


def main():
    input_dir = INPUT_DIR
    if not input_dir:
        print("未设置 PROCESS_INPUT_DIR 环境变量")
        return

    if not os.path.exists(input_dir):
        print(f"输入目录不存在: {input_dir}")
        return

    # 已精减子目录（在已标注目录下）
    jingjian_dir = os.path.join(input_dir, "已精减")

    # IBKR 输出目录
    ibkr_dir = IBKR_OUTPUT_DIR

    # 写入前删除已有的已精减文件
    print(f"\n清理已精减目录: {jingjian_dir}")
    clear_jingjian_files(jingjian_dir)

    print(f"\n清理IBKR目录: {ibkr_dir}")
    clear_jingjian_files(ibkr_dir)

    files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    files = [f for f in files if OUTPUT_SUFFIX not in os.path.basename(f)]
    if not files:
        print("当前目录没有找到 .xlsx 文件")
        return
    print(f"共找到 {len(files)} 个待处理文件")
    for f in files:
        try:
            process_file(f, jingjian_dir, ibkr_dir)
        except Exception as e:
            import traceback
            print(f"  处理失败: {f}")
            traceback.print_exc()
    print("\n全部处理完成！")


if __name__ == "__main__":
    main()
