"""
股票价格区间双匹配处理脚本 V5.1
功能：
1. 读取源Excel文件
2. Sheet1（价格区间处理结果）：新增AK列（双区间匹配），按AK列倒序排序
3. Sheet2（双区间再分表格）：保留原有结构，处理所有AK列不为空的股票
   - A列（股票代码）：股票代码
   - B列（现价）：最新价格
   - C列（最高）：对应区间左边界列的值
   - D列（最低）：对应区间右边界列的值
   - E-AK列：保留原有公式，行数不足时自动向下复制公式
4. 兼容：价格列名支持「现价」和「最新价格」
5. 兼容：第二个Sheet名支持「Sheet1」、「双区间再分表格」等
6. 内置模板：输入文件没有第二个Sheet时，自动用内置模板创建
7. 修正：1-2区间时，C列（最高）取历史最高价格列的值
8. 修正：修复公式模板中的循环引用问题

使用方法：
    python stock_price_range_processor_v5.1.py <输入文件路径> <输出文件路径>
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
import sys
import os
import re
from pathlib import Path

# ===================== 配置区（只改这里） =====================
# 市场选择：US / A / HK（改这一个变量即可切换市场，或通过环境变量 MARKET 传入）
MARKET = os.environ.get("MARKET", "US")

# 基础路径（自动根据 MARKET 生成输入/输出目录）
BASE_DIR = r"/home/ubuntu/alibaba_agent/downlaods_futu_data/data_all_market"
INPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo", "计算区间")                    # 输入目录
OUTPUT_DIR = os.path.join(BASE_DIR, MARKET, "demo", "计算区间", "双区间相同再算幅度")  # 输出目录
# ============================================================


# ============================================================
# Sheet2 内置模板定义（修复循环引用）
# ============================================================
SHEET2_TEMPLATE = {
    'sheet_name': '双区间再分表格',
    
    # 表头定义（37列）
    'headers': [
        '股票代码',
        '现价',
        '最高(双8-16)',
        '最低(双8-16)',
        '双16再分高位中间',
        '双16再分中间',
        '双16再分低位中间',
        '双16再分次1',
        '双16再分次2',
        '双16再分次3',
        '双16再分次4',
        '双16再分次5',
        '双16再分次6',
        '双16再分次7',
        '双16再分次8',
        '双16再分次9',
        '双16再分次10',
        '双16再分次11',
        '双16再分次12',
        '双16再分次13',
        '双16再分次14',
        '双16再分次15',
        '双16再分次16',
        '双16再分次17',
        '双16再分次18',
        '双16再分次19',
        '双16再分次20',
        '双16再分次21',
        '双16再分次22',
        '双16再分次23',
        '双16再分次24',
        '双16再分次25',
        '双16再分次26',
        '双16再分次27',
        '双16再分次28',
        '双16再分次29',
        '双16再分次30',
    ],
    
    # 第2行公式定义（列索引从0开始，即A=0, B=1, ...）
    # E列到AK列（索引4到36）
    # 公式逻辑：
    #   E列 = (C2 + F2) / 2     (高位中间)
    #   F列 = (C2 + D2) / 2     (中间)
    #   G列 = (D2 + F2) / 2     (低位中间)
    #   H列 = (G2 + D2) / 2     (次1)
    #   I列 = (H2 + D2) / 2     (次2)
    #   ... 以此类推，每列取前一列与D列的平均值
    #   AJ列 = (AI2 + D2) / 2   (次29)
    #   AK列 = (AJ2 + D2) / 2   (次30)
    'formulas_row2': [
        None,  # A列：股票代码
        None,  # B列：现价
        None,  # C列：最高(双8-16)
        None,  # D列：最低(双8-16)
        '=SUM(C2+F2)/2',      # E列：双16再分高位中间
        '=SUM(C2+D2)/2',      # F列：双16再分中间
        '=SUM(D2+F2)/2',      # G列：双16再分低位中间
        '=(G2+D2)/2',         # H列：双16再分次1
        '=(H2+D2)/2',         # I列：双16再分次2
        '=(I2+D2)/2',         # J列：双16再分次3
        '=(J2+D2)/2',         # K列：双16再分次4
        '=(K2+D2)/2',         # L列：双16再分次5
        '=(L2+D2)/2',         # M列：双16再分次6
        '=(M2+D2)/2',         # N列：双16再分次7
        '=(N2+D2)/2',         # O列：双16再分次8
        '=(O2+D2)/2',         # P列：双16再分次9
        '=(P2+D2)/2',         # Q列：双16再分次10
        '=(Q2+D2)/2',         # R列：双16再分次11
        '=(R2+D2)/2',         # S列：双16再分次12
        '=(S2+D2)/2',         # T列：双16再分次13
        '=(T2+D2)/2',         # U列：双16再分次14  <-- 修复：原模板此处缺失
        '=(U2+D2)/2',         # V列：双16再分次15  <-- 修复：原模板此处缺失
        '=(V2+D2)/2',         # W列：双16再分次16
        '=(W2+D2)/2',         # X列：双16再分次17
        '=(X2+D2)/2',         # Y列：双16再分次18
        '=(Y2+D2)/2',         # Z列：双16再分次19
        '=(Z2+D2)/2',         # AA列：双16再分次20
        '=(AA2+D2)/2',        # AB列：双16再分次21
        '=(AB2+D2)/2',        # AC列：双16再分次22
        '=(AC2+D2)/2',        # AD列：双16再分次23
        '=(AD2+D2)/2',        # AE列：双16再分次24
        '=(AE2+D2)/2',        # AF列：双16再分次25
        '=(AF2+D2)/2',        # AG列：双16再分次26
        '=(AG2+D2)/2',        # AH列：双16再分次27
        '=(AH2+D2)/2',        # AI列：双16再分次28
        '=(AI2+D2)/2',        # AJ列：双16再分次29
        '=(AJ2+D2)/2',        # AK列：双16再分次30
    ],
    
    # 默认数据行数（含表头共34行）
    'default_rows': 34,
}


def parse_range_left(range_str):
    """
    解析区间字符串的左边界数值，用于排序
    """
    if pd.isna(range_str) or range_str == '' or range_str is None:
        return -1
    if isinstance(range_str, str) and range_str.startswith('≤'):
        return 1
    if isinstance(range_str, str) and '-' in range_str:
        try:
            left = int(range_str.split('-')[0])
            return left
        except (ValueError, IndexError):
            return 0
    return 0


def get_range_columns(range_str):
    """
    根据区间字符串获取对应的最高列名和最低列名
    特殊处理：1-2区间返回特殊标记
    """
    if pd.isna(range_str) or range_str == '' or range_str is None:
        return None, None
    
    if isinstance(range_str, str) and range_str.startswith('≤'):
        val = range_str.replace('≤', '').strip()
        return val, None
    
    if isinstance(range_str, str) and '-' in range_str:
        parts = range_str.split('-')
        if len(parts) == 2:
            left = parts[0].strip()
            right = parts[1].strip()
            
            # 特殊处理1-2区间：最高列使用历史最高价格
            if left == '1' and right == '2':
                return 'HISTORICAL_HIGH', right
            
            return left, right
    
    return None, None


def get_price_column(df):
    """
    获取价格列名（兼容「现价」和「最新价格」）
    """
    if '现价' in df.columns:
        return '现价'
    elif '最新价格' in df.columns:
        return '最新价格'
    else:
        raise ValueError('源文件中未找到价格列（支持「现价」或「最新价格」）')


def get_stock_code_column(df):
    """
    获取股票代码列名（兼容「股票代码」和「代码」等）
    """
    if '股票代码' in df.columns:
        return '股票代码'
    elif '代码' in df.columns:
        return '代码'
    else:
        raise ValueError('源文件中未找到股票代码列')


def get_historical_high_column(df):
    """
    获取历史最高价格列名
    """
    if '历史最高价格' in df.columns:
        return '历史最高价格'
    elif '历史最高价' in df.columns:
        return '历史最高价'
    else:
        return None


def find_sheet2(wb):
    """
    查找第二个Sheet（兼容不同名称）
    返回：(sheet对象, sheet名称)
    如果找不到，返回(None, None)
    """
    # 可能的Sheet2名称
    possible_names = ['Sheet1', '双区间再分表格', 'Sheet2', 'sheet2']
    
    for name in possible_names:
        if name in wb.sheetnames:
            return wb[name], name
    
    # 如果有第二个sheet，就用第二个
    if len(wb.sheetnames) >= 2:
        name = wb.sheetnames[1]
        return wb[name], name
    
    return None, None


def create_sheet2_from_template(wb):
    """
    使用内置模板创建Sheet2（修复循环引用问题）
    """
    template = SHEET2_TEMPLATE
    ws = wb.create_sheet(title=template['sheet_name'])
    
    # 写入表头
    for col_idx, header in enumerate(template['headers']):
        ws.cell(row=1, column=col_idx + 1, value=header)
    
    # 写入第2行公式
    for col_idx, formula in enumerate(template['formulas_row2']):
        if formula:
            ws.cell(row=2, column=col_idx + 1, value=formula)
    
    # 复制公式到默认行数
    default_rows = template['default_rows']
    for row in range(3, default_rows + 1):
        for col_idx, formula_template in enumerate(template['formulas_row2']):
            if formula_template:
                # 使用正则表达式替换行号
                def replace_row_num(match):
                    col_letter = match.group(1)
                    return f'{col_letter}{row}'
                new_formula = re.sub(r'([A-Z]+)2', replace_row_num, formula_template)
                ws.cell(row=row, column=col_idx + 1, value=new_formula)
    
    print(f"  已使用内置模板创建Sheet2：{template['sheet_name']}")
    print(f"  共 {default_rows} 行，{len(template['headers'])} 列")
    print(f"  公式已修复，无循环引用")
    
    return ws, template['sheet_name']


def copy_formula_row(ws, source_row, target_row, max_col):
    """
    复制一行的公式到目标行
    """
    for col in range(1, max_col + 1):
        source_cell = ws.cell(row=source_row, column=col)
        target_cell = ws.cell(row=target_row, column=col)
        
        if source_cell.value and isinstance(source_cell.value, str) and source_cell.value.startswith('='):
            # 复制公式，替换行号
            formula = source_cell.value
            
            def replace_row(match):
                col_letter = match.group(1)
                return f'{col_letter}{target_row}'
            
            new_formula = re.sub(r'([A-Z]+)\d+', replace_row, formula)
            target_cell.value = new_formula
        
        # 复制样式
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = copy(source_cell.number_format)
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)


def repair_existing_sheet2_formulas(ws):
    """
    修复现有Sheet2中的公式循环引用问题
    """
    print("  正在检查并修复现有Sheet2的公式...")
    
    repaired_count = 0
    
    # 检查数据行数
    max_row = ws.max_row
    
    # 对每一行数据（从第2行开始）
    for row in range(2, max_row + 1):
        # 检查E列到AK列（第5列到第37列）
        for col in range(5, 38):  # E=5, AK=37
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                # 检查公式中是否引用了当前行
                formula = cell.value
                # 检查循环引用：公式中是否包含当前行的引用
                # 例如：AJ17 的公式中如果引用了 AJ17 本身，就是循环引用
                col_letter = get_column_letter(col)
                ref_pattern = f'{col_letter}{row}'
                
                # 如果公式包含对自身的引用，说明有循环引用，需要修复
                if ref_pattern in formula:
                    # 对于次列公式，应该是引用前一列的值
                    # 例如：AJ列的公式应该是 =(AI{row}+D{row})/2
                    if col > 5:  # 不是E列
                        prev_col = col - 1
                        prev_col_letter = get_column_letter(prev_col)
                        new_formula = f'=({prev_col_letter}{row}+D{row})/2'
                        cell.value = new_formula
                        repaired_count += 1
                        # print(f"    修复第{row}行第{col_letter}列: {formula} -> {new_formula}")
    
    print(f"  已修复 {repaired_count} 个循环引用公式")
    return repaired_count


def process_stock_data(input_file, output_file):
    """
    主处理函数
    """
    print(f"正在读取源文件: {input_file}")
    
    # ============================================================
    # 第一步：用pandas读取sheet1数据，进行计算和排序
    # ============================================================
    df_sheet1 = pd.read_excel(input_file, sheet_name=0)
    sheet1_name = pd.ExcelFile(input_file).sheet_names[0]
    print(f"  Sheet1名称: {sheet1_name}")
    print(f"  读取到 {len(df_sheet1)} 行数据，{len(df_sheet1.columns)} 列")
    
    # 获取价格列名
    price_col = get_price_column(df_sheet1)
    print(f"  价格列名: {price_col}")
    
    # 获取股票代码列名
    code_col = get_stock_code_column(df_sheet1)
    print(f"  股票代码列名: {code_col}")
    
    # 获取历史最高价格列名
    high_price_col = get_historical_high_column(df_sheet1)
    if high_price_col:
        print(f"  历史最高价格列名: {high_price_col}")
    else:
        print(f"  ⚠ 未找到历史最高价格列，1-2区间将无法正确填充")
    
    # 检查必要的列
    required_cols = ['最新价格区间', '历史最低区间']
    for col in required_cols:
        if col not in df_sheet1.columns:
            raise ValueError(f"源文件Sheet1中缺少必要列: {col}")
    
    # 新增双区间匹配列（AK列）
    print("\n正在计算双区间匹配列（AK列）...")
    df_sheet1['双区间匹配'] = ''
    match_mask = df_sheet1['最新价格区间'] == df_sheet1['历史最低区间']
    df_sheet1.loc[match_mask, '双区间匹配'] = df_sheet1.loc[match_mask, '最新价格区间']
    
    matched_count = match_mask.sum()
    print(f"  双区间匹配总行数: {matched_count}")
    
    # 按双区间匹配列倒序排序
    print("\n正在按双区间匹配列倒序排序...")
    df_sheet1['_sort_key'] = df_sheet1['双区间匹配'].apply(parse_range_left)
    df_sheet1 = df_sheet1.sort_values('_sort_key', ascending=False, kind='mergesort')
    df_sheet1 = df_sheet1.drop('_sort_key', axis=1)
    df_sheet1 = df_sheet1.reset_index(drop=True)
    print("  排序完成")
    
    # 提取所有双区间匹配的股票
    matched_stocks = df_sheet1[df_sheet1['双区间匹配'] != ''].copy()
    matched_stocks = matched_stocks.reset_index(drop=True)
    
    print(f"\n双区间匹配股票（共{len(matched_stocks)}只，按倒序排列）:")
    range_counts = matched_stocks['双区间匹配'].value_counts()
    for range_str in sorted(range_counts.index, key=parse_range_left, reverse=True):
        count = range_counts[range_str]
        high_col, low_col = get_range_columns(range_str)
        if high_col == 'HISTORICAL_HIGH':
            high_col_desc = f'历史最高价格(特殊处理)'
        else:
            high_col_desc = high_col
        print(f"  {range_str}: {count} 只  (最高列={high_col_desc}, 最低列={low_col})")
    
    # ============================================================
    # 第二步：用openpyxl加载工作簿
    # ============================================================
    print("\n正在加载工作簿...")
    wb = load_workbook(input_file)
    
    # 获取第一个sheet
    ws_price = wb[sheet1_name]
    
    # 查找或创建第二个sheet
    ws_sheet2, sheet2_name = find_sheet2(wb)
    if ws_sheet2 is None:
        print("  未找到Sheet2，正在使用内置模板创建...")
        ws_sheet2, sheet2_name = create_sheet2_from_template(wb)
    else:
        print(f"  Sheet2名称: {sheet2_name}")
        
        # 修复现有Sheet2的循环引用问题
        repair_existing_sheet2_formulas(ws_sheet2)
    
    original_rows = ws_sheet2.max_row
    original_cols = ws_sheet2.max_column
    print(f"  Sheet2原有行数: {original_rows}")
    print(f"  Sheet2原有列数: {original_cols}")
    
    # ============================================================
    # 第三步：更新Sheet1（价格区间处理结果）
    # ============================================================
    print("\n正在更新Sheet1（价格区间处理结果）...")
    
    # 在AK列（第37列）写入表头
    ws_price.cell(row=1, column=37, value='双区间匹配')
    
    # 重写所有数据行
    for row_idx in range(len(df_sheet1)):
        excel_row = row_idx + 2
        for col_idx in range(len(df_sheet1.columns)):
            excel_col = col_idx + 1
            value = df_sheet1.iloc[row_idx, col_idx]
            
            if pd.isna(value):
                value = None
            
            cell = ws_price.cell(row=excel_row, column=excel_col)
            cell.value = value
    
    print(f"  Sheet1更新完成，共 {len(df_sheet1)} 行数据")
    
    # ============================================================
    # 第四步：扩展Sheet2的公式行（如果需要）
    # ============================================================
    needed_rows = len(matched_stocks) + 1  # 1行表头 + 数据行
    print(f"\n需要的数据行数: {len(matched_stocks)}")
    print(f"原有数据行数: {original_rows - 1}")
    
    if needed_rows > original_rows:
        print(f"行数不足，正在向下复制公式到第{needed_rows}行...")
        
        template_row = 2
        for target_row in range(original_rows + 1, needed_rows + 1):
            copy_formula_row(ws_sheet2, template_row, target_row, original_cols)
        
        print(f"  已扩展到 {needed_rows} 行")
    else:
        print("  行数充足，无需扩展")
    
    # ============================================================
    # 第五步：填充Sheet2的数据（修正1-2区间处理）
    # ============================================================
    print("\n正在填充Sheet2数据...")
    
    start_row = 2
    special_range_count = 0
    
    for i, row in matched_stocks.iterrows():
        excel_row = start_row + i
        range_str = row['双区间匹配']
        high_col, low_col = get_range_columns(range_str)
        
        # A列：股票代码
        ws_sheet2.cell(row=excel_row, column=1, value=row[code_col])
        
        # B列：现价
        ws_sheet2.cell(row=excel_row, column=2, value=row[price_col])
        
        # C列：最高值
        if high_col == 'HISTORICAL_HIGH':
            # 1-2区间特殊处理：使用历史最高价格
            if high_price_col and high_price_col in df_sheet1.columns:
                val = row[high_price_col]
                ws_sheet2.cell(row=excel_row, column=3, value=val)
                special_range_count += 1
            else:
                print(f"    ⚠ 警告: 1-2区间股票 {row[code_col]} 无法获取历史最高价格")
        elif high_col and high_col in df_sheet1.columns:
            ws_sheet2.cell(row=excel_row, column=3, value=row[high_col])
        else:
            print(f"    ⚠ 警告: 无法获取 {row[code_col]} 的最高值列 {high_col}")
        
        # D列：最低值
        if low_col and low_col in df_sheet1.columns:
            ws_sheet2.cell(row=excel_row, column=4, value=row[low_col])
    
    print(f"  已填充 {len(matched_stocks)} 行数据（第{start_row}行到第{start_row + len(matched_stocks) - 1}行）")
    print(f"  其中1-2区间特殊处理: {special_range_count} 只")
    print("  E列及以后的公式列保持不变，将自动计算")
    
    # ============================================================
    # 第六步：保存到输出文件
    # ============================================================
    print(f"\n正在保存结果到: {output_file}")
    
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    wb.save(output_file)
    print("  保存成功！")
    
    # ============================================================
    # 输出统计信息
    # ============================================================
    print("\n" + "="*60)
    print("处理完成！统计信息：")
    print("="*60)
    print(f"总股票数: {len(df_sheet1)}")
    print(f"双区间匹配总数: {matched_count}")
    print(f"\n各区间匹配分布:")
    for range_str in sorted(range_counts.index, key=parse_range_left, reverse=True):
        count = range_counts[range_str]
        high_col, low_col = get_range_columns(range_str)
        if high_col == 'HISTORICAL_HIGH':
            high_col_desc = f'历史最高价格(特殊处理)'
        else:
            high_col_desc = high_col
        print(f"  {range_str}: {count} 只  (最高列={high_col_desc}, 最低列={low_col})")
    
    if special_range_count > 0:
        print(f"\n⚠ 1-2区间特殊处理: {special_range_count} 只")
        print(f"  这些股票的C列(最高)使用了历史最高价格列的值")
    
    print(f"\nSheet2信息:")
    print(f"  Sheet名称: {sheet2_name}")
    print(f"  数据行数: {len(matched_stocks)}")
    print(f"  总列数: {original_cols}")
    print(f"  A列（股票代码）：股票代码")
    print(f"  B列（现价）：{price_col}")
    print(f"  C列（最高）：对应区间左边界列的值（1-2区间特殊处理为历史最高价格）")
    print(f"  D列（最低）：对应区间右边界列的值")
    print(f"  E列及以后：保留原有公式，自动计算细分档位")
    print(f"\n兼容性说明:")
    print(f"  ✓ 价格列名兼容「现价」和「最新价格」")
    print(f"  ✓ Sheet2名称兼容多种命名")
    print(f"  ✓ 无Sheet2时自动使用内置模板创建")
    print(f"  ✓ 1-2区间特殊处理：C列使用历史最高价格")
    print(f"  ✓ 修复了公式循环引用问题")
    print(f"\n输出文件: {output_file}")
    print("="*60)
def get_file_list(folder):
    p = Path(folder)
    file_names = [file.name for file in p.iterdir() if file.is_file() and file.suffix.lower() in (".xlsx", ".xls")]
    return file_names


def main():
    """
    命令行入口
    """
    if len(sys.argv) < 3:
        print("使用方法:")
        print("  python stock_price_range_processor_v5.1.py <输入文件路径> <输出文件路径>")
        print("\n示例:")
        print("  python stock_price_range_processor_v5.1.py 股票价格区间处理结果.xlsx 处理结果_v5.1.xlsx")
        print("\n或者直接运行使用默认文件:")
        print("  python stock_price_range_processor_v5.1.py")
        print("  （默认读取当前目录下的 股票价格区间处理结果.xlsx，输出到 处理结果_v5.1.xlsx）")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    
    if not os.path.exists(input_file):
        print(f"错误：输入文件不存在: {input_file}")
        sys.exit(1)
    
    try:
        process_stock_data(input_file, output_file)
    except Exception as e:
        print(f"\n处理出错: {e}")
        import traceback    
        traceback.print_exc()
        sys.exit(1)

#""
# soureFile = r"D:\D盘存储目录\股票数据\所有USHKA不足3年3-6年及以上\A股\1970-2017".strip()


if __name__ == '__main__':
    if len(sys.argv) == 1:          
        dir_path = INPUT_DIR
        output_dir = OUTPUT_DIR
        
        # 如果输出目录不存在则创建
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"📂 输入目录: {dir_path}")
        print(f"📂 输出目录: {output_dir}")
        
        res = get_file_list(dir_path)
        print(f"检测到待处理Excel文件共 {len(res)} 个：{res}\n")
        for file in res:
            try:
                input_file = os.path.join(dir_path, file)  
                # 输出文件名：原文件名（去掉.xlsx）+ 双区间相同再算幅度.xlsx
                stem = Path(file).stem
                output_file = os.path.join(output_dir, f"{stem}双区间相同再算幅度.xlsx")
                process_stock_data(input_file, output_file)
            except Exception as err:
                print(f"⚠️ 文件 {file} 处理异常跳过：{err}\n")






    #     Input_File = r'\20260820001443所有A股导出1998-2005_592支股票(6年以上)_2双区间相同再计算幅度区间_v6.1.xlsx'.strip()
    #     output_File = str(Input_File[:Input_File.find("demo")])+ '双区间相同再计算幅度区间_v6.1' + ".xlsx"

    #     default_input = soureFile + Input_File
    #     default_output = soureFile + output_File 

    #     if os.path.exists(default_input):
    #         print(f"使用默认文件: 输入={default_input}, 输出={default_output}")
    #         process_stock_data(default_input, default_output)
    #     else:
    #         main()
    # else:
    #     main()
