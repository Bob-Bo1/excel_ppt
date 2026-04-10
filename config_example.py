# -*- coding: utf-8 -*-
"""
配置文件示例生成脚本
运行此脚本可生成示例 config.xlsx 文件
"""

import pandas as pd

# 创建示例配置数据
data = {
    "页面编号": [1, 1, 2],
    "页面标题": ["销售概览", "销售概览", "利润分析"],
    "透视行": ["月份", "月份", "月份"],
    "透视列": ["区域", "", "区域"],
    "透视值": ["销售额", "销售额,利润", "利润"],
    "计算方式": ["sum", "sum", "mean"],
    "图表类型": ["柱形图", "组合图", "折线图"],
    "柱形列": ["", "销售额", ""],
    "折线列": ["", "利润", ""]
}

# 创建DataFrame
df = pd.DataFrame(data)

# 保存为Excel
output_path = "config_example.xlsx"
df.to_excel(output_path, index=False)

print(f"示例配置文件已生成：{output_path}")
print("\n配置说明：")
print("1. 页面编号：同一编号的配置会生成在同一页PPT")
print("2. 页面标题：PPT页面标题")
print("3. 透视行：X轴分组维度")
print("4. 透视列：系列展开维度（可选）")
print("5. 透视值：要统计的数字字段，多列用逗号分隔")
print("6. 计算方式：sum/mean/count")
print("7. 图表类型：柱形图/折线图/组合图")
print("8. 柱形列/折线列：仅组合图需要，指定哪些字段显示为柱形或折线")
