import openpyxl
wb = openpyxl.load_workbook('損傷リスト.xlsx')
ws = wb["Sheet1"]#Sheet1を読み込む

wb1 = openpyxl.load_workbook('状況表.xlsx')
ws1 = wb1["４月"]#Sheet1を読み込む


count_map = {}
for row in ws.iter_rows(min_row=5):
    所属 = row[3].value 
    if 所属 not in count_map:
        count_map[所属] = 0
    count_map[所属] += 1
print(count_map)

affiliation_map = {}
for row in ws1.iter_rows(min_row=19):
    所属1 = row[1].value
    print(所属1)

if 所属 == 所属1:
    num = count_map['装備']
print(num)    

