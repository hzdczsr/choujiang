import pandas as pd
import json

try:
    # 读取Excel文件
    df = pd.read_excel('d:\\原C盘\\download\\choujiang\\namexlsx.xlsx')
    
    print("Excel文件读取成功！")
    print(f"数据形状: {df.shape}")
    print(f"列名: {df.columns.tolist()}")
    print("\n前几行数据:")
    print(df.head())
    
    # 尝试提取姓名数据
    names = []
    
    # 检查所有列，寻找姓名数据
    for col in df.columns:
        print(f"\n检查列 '{col}':")
        print(df[col].dropna().head(10).tolist())
        
        # 尝试将这一列作为姓名
        col_names = df[col].dropna().astype(str).tolist()
        if col_names:
            # 如果这一列有数据，将其添加到姓名列表
            names.extend(col_names)
    
    # 去重并过滤空值
    names = list(set([name for name in names if name and str(name).strip() and str(name) != 'nan']))
    
    print(f"\n提取到的姓名（共{len(names)}个）:")
    for i, name in enumerate(names, 1):
        print(f"{i}. {name}")
    
    # 保存到JSON文件
    with open('names.json', 'w', encoding='utf-8') as f:
        json.dump(names, f, ensure_ascii=False, indent=2)
    
    print(f"\n姓名数据已保存到 names.json")
    
except Exception as e:
    print(f"读取Excel文件时出错: {e}")
    import traceback
    traceback.print_exc()