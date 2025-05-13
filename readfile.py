
# 文件名
filename = 'D:\example.txt'

# 打开文件
with open(filename, 'r', encoding='utf-8') as file:
    # 按行读取文件内容并输出
    for line in file:
        print(line.strip())  # 使用 strip() 去掉每行末尾的换行符

# 提示文件已读取完毕
print(f"文件 {filename} 已读取完毕。")