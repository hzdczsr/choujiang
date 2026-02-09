#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import json
import sys

def analyze_excel_file():
    """分析Excel文件的详细内容"""
    excel_file = 'namexlsx.xlsx'
    
    try:
        print(f"正在分析Excel文件: {excel_file}")
        print("=" * 50)
        
        # 使用pandas读取Excel文件
        excel_file_obj = pd.ExcelFile(excel_file)
        sheet_names = excel_file_obj.sheet_names
        
        print(f"工作表数量: {len(sheet_names)}")
        print(f"工作表名称: {sheet_names}")
        print("=" * 50)
        
        all_data = {}
        total_cells = 0
        non_empty_cells = 0
        valid_names = []
        
        # 遍历所有工作表
        for sheet_name in sheet_names:
            print(f"\n分析工作表: {sheet_name}")
            print("-" * 30)
            
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            sheet_data = []
            
            print(f"工作表大小: {df.shape[0]} 行 x {df.shape[1]} 列")
            
            # 遍历所有单元格
            for row_idx in range(len(df)):
                for col_idx in range(len(df.columns)):
                    total_cells += 1
                    cell_value = df.iloc[row_idx, col_idx]
                    
                    if pd.notna(cell_value) and str(cell_value).strip():
                        non_empty_cells += 1
                        cell_str = str(cell_value).strip()
                        sheet_data.append({
                            'row': row_idx + 1,
                            'col': col_idx + 1,
                            'value': cell_str,
                            'length': len(cell_str)
                        })
                        
                        # 检查是否为可能的姓名
                        if is_likely_name(cell_str):
                            valid_names.append(cell_str)
                            print(f"  姓名候选: 行{row_idx+1}, 列{col_idx+1} = '{cell_str}'")
            
            all_data[sheet_name] = {
                'data': sheet_data,
                'valid_names': list(set([name for name in valid_names if name in sheet_data]))
            }
        
        print("\n" + "=" * 50)
        print("统计信息:")
        print(f"总单元格数: {total_cells}")
        print(f"非空单元格数: {non_empty_cells}")
        print(f"有效姓名数量: {len(set(valid_names))}")
        print(f"有效姓名列表: {sorted(set(valid_names))}")
        
        # 保存分析结果到JSON文件
        analysis_result = {
            'summary': {
                'total_cells': total_cells,
                'non_empty_cells': non_empty_cells,
                'sheet_count': len(sheet_names),
                'sheet_names': sheet_names
            },
            'sheets': all_data,
            'valid_names': sorted(list(set(valid_names)))
        }
        
        with open('excel_analysis.json', 'w', encoding='utf-8') as f:
            json.dump(analysis_result, f, ensure_ascii=False, indent=2)
        
        print(f"\n详细分析结果已保存到: excel_analysis.json")
        
    except Exception as e:
        print(f"分析失败: {str(e)}")
        import traceback
        traceback.print_exc()

def is_likely_name(text):
    """判断文本是否可能是姓名"""
    if not text or len(text.strip()) == 0:
        return False
    
    text = text.strip()
    
    # 排除明显不是姓名的情况
    if any(keyword in text.lower() for keyword in [
        '姓名', '名字', 'name', 'sheet', '工作表', '页', '表',
        'nan', 'null', 'undefined', 'error', '错误',
        '1', '2', '3', '4', '5', '6', '7', '8', '9', '0',
        'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
        'test', '测试', '示例', 'sample'
    ]):
        return False
    
    # 中文姓名通常2-4个字符，且不包含特殊符号
    if len(text) >= 2 and len(text) <= 6:
        # 检查是否只包含中文字符、常见符号和英文字母
        import re
        # 允许中文、常见的英文字母、数字（可能是编号）
        if re.match(r'^[一-龥a-zA-Z0-9\(\)\[\]【】（）【】\.\-_\s]+$', text):
            return True
    
    return False

if __name__ == "__main__":
    analyze_excel_file()