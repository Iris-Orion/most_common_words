import openpyxl
from collections import Counter

# 打开Excel文件
wb = openpyxl.load_workbook('file1.xlsx')
ws = wb.active

# 提取所有单元格中的日文文本
texts = []
for row in ws.iter_rows():
    for cell in row:
        if cell.value and isinstance(cell.value, str):
            texts.append(cell.value)

#统计日文文本出现次数  
text_counts = Counter(texts)

# # 获取出现次数最多的日文文本
# most_common_text, max_count = text_counts.most_common(1)[0]

# print(f'出现次数最多的日文文本是:{most_common_text},出现了{max_count}次')

# 打印每个文本出现的次数
print("文本 - 出现次数")
for text in set(texts):
    print(f"{text} - {text_counts[text]}")