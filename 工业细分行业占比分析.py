import pandas as pd
import json

# 读取Excel文件
excel_file = '/Users/zhangpingzuan/PycharmProjects/finance-tool/data/2020.xlsx'  # 替换为你的Excel文件路径

def extract_columns_from_excel(file_path, start_row, end_row , column_names):
    # 读取Excel文件，从指定行开始读取
    df = pd.read_excel(file_path, skiprows=start_row - 1, nrows=end_row - start_row + 1)

    # 提取指定的列
    extracted_columns = df[column_names]


    # 按照"GO"列的数值排序
    sorted_result = extracted_columns.sort_values(by='GO')

    # 将提取的列数据转换为JSON格式
    json_result = sorted_result.to_json(orient='records', force_ascii=False)

    return json_result


# 示例用法
start_row = 6  # 替换为实际的开始行号
end_row = 160  # 替换为实际的开始行号
start_column = '部门名称'  # 替换为实际的开始列名
end_column = '代码'      # 替换为实际的结束列名
column_names = ['部门名称', '代码', 'GO']  # 替换为实际的列名列表
output_file = 'output.json'  # 替换为实际的输出文件路径

result = extract_columns_from_excel(excel_file, start_row, end_row, column_names)

# 将JSON数据写入文件
with open(output_file, 'w') as f:
    f.write(result)
