import pandas as pd

# Excel 文件名
filename = 'E:\python_workspace\demo\example.xlsx'

# 使用 pandas 读取 Excel 文件
df = pd.read_excel(filename, engine='openpyxl')
#获取行列的值
print(df.iloc[1,2])
# 按行遍历 DataFrame 并输出每行的内容
for index, row in df.iterrows():
    print(row.values)  # 输出每行的值，作为数组
    # 如果你想要更格式化的输出，可以这样做：
    # print(f"Row {index + 1}: {', '.join(map(str, row.values))}")

# 提示文件已读取完毕
print(f"文件 {filename} 已读取完毕。")
