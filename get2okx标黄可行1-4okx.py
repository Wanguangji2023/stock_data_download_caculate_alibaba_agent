"""
双区间再分表格 - 现价区间上下单元格标黄工具
依赖：pip install openpyxl
"""
import openpyxl
from openpyxl.styles import PatternFill
from pathlib import Path
import sys
import os

# ===================== 配置区（只改这里） =====================
# 市场选择：US / A / HK（改这一个变量即可切换市场，或通过环境变量 MARKET 传入）
MARKET = os.environ.get("MARKET", "US")

# 基础路径（自动根据 MARKET 生成输入/输出目录）
BASE_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market"
INPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo", "计算区间", "双区间相同再算幅度")           # 输入目录
OUTPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo", "计算区间", "双区间相同再算幅度", "标黄")  # 输出目录
# ============================================================

def get_file_list(folder):
    p = Path(folder)
    file_names = [file.name for file in p.iterdir() if file.is_file() and file.suffix.lower() in (".xlsx", ".xls")]
    return file_names
def calc_interval_values(c_val, d_val):
    """根据最高(C)、最低(D)计算E~AK所有区间数值，返回{列序号:数值}"""
    val_map = {}
    # E=5 F=6 G=7 H=8 ... AK=37
    f_mid = (c_val + d_val) / 2
    val_map[6] = f_mid                # F列
    val_map[5] = (c_val + f_mid) / 2  # E列
    g_mid = (d_val + f_mid) / 2
    val_map[7] = g_mid                # G列
    cur = g_mid
    # H(8) ~ AK(37) 循环递推
    for col_idx in range(8, 38):
        cur = (cur + d_val) / 2
        val_map[col_idx] = cur
    return val_map

def get_upper_lower(val_dict, price):
    """获取小于现价最大值(下界)、大于现价最小值(上界)"""
    # 按数值从小到大排序
    sorted_list = sorted(val_dict.items(), key=lambda x: x[1])
    lower_col = None
    upper_col = None
    for col, v in sorted_list:
        if v == price:
            # 价格刚好等于区间值，只标记当前单元格
            return col, None
        elif v < price:
            lower_col = col
        else:
            upper_col = col
            break
    return lower_col, upper_col

def excel_mark_yellow(input_file, output_file):
    # 打开Excel，保留原有格式
    wb = openpyxl.load_workbook(input_file)
    sheet_name = "双区间再分表格"
    ws = wb[sheet_name]
    # 黄色填充样式
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 从第2行开始循环（第1行表头）
    for row_num in range(2, ws.max_row + 1):
        # 获取单元格数值
        price_cell = ws.cell(row=row_num, column=2)  # B列 现价
        high_cell = ws.cell(row=row_num, column=3)   # C列 最高(双8-16)
        low_cell = ws.cell(row=row_num, column=4)    # D列 最低(双8-16)

        price = price_cell.value
        high = high_cell.value
        low = low_cell.value

        # 空数据跳过，防止计算报错
        if price is None or high is None or low is None:
            continue
        if not isinstance(price, (int, float)) or not isinstance(high, (int, float)) or not isinstance(low, (int, float)):
            continue

        # 计算所有区间值
        interval_data = calc_interval_values(high, low)
        lower_col, upper_col = get_upper_lower(interval_data, price)

        # 填充黄色
        if lower_col is not None:
            ws.cell(row=row_num, column=lower_col).fill = yellow_fill
        if upper_col is not None:
            ws.cell(row=row_num, column=upper_col).fill = yellow_fill

    # 保存新文件
    wb.save(output_file)
    wb.close()
    print(f"处理完成！输出文件：{output_file}")
#专用于处理OKX数据的标黄
if __name__ == "__main__":
    filepath = INPUT_DIR
    output_dir = OUTPUT_DIR
    
    # 如果输出目录不存在则创建
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"📂 输入目录: {filepath}")
    print(f"📂 输出目录: {output_dir}")
    
    res = get_file_list(filepath)
    print(f"检测到待处理Excel文件共 {len(res)} 个：{res}\n")
    for file in res:
        try:    
            input_file = os.path.join(filepath, file)
            # 输出文件名：原文件名（去掉.xlsx）+ 幅度计算okx标黄.xlsx
            stem = Path(file).stem
            output_file = os.path.join(output_dir, f"{stem}幅度计算okx标黄.xlsx")
            excel_mark_yellow(input_file, output_file)
        except Exception as err:
            print(f"⚠️ 文件 {file} 处理异常跳过：{err}\n")
        
