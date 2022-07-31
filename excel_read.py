import openpyxl
wb = openpyxl.load_workbook('損傷リスト.xlsx')
ws = wb["Sheet1"]#Sheet1を読み込む




syozoku_map = {}
for row in ws.iter_rows(min_row=5):
    所属 = row[3].value
    if 所属 == None:
        continue
    金額 = row[10].value 
    if 所属 not in syozoku_map:
        syozoku_map[所属] = {
            'count':0,
            'money':0
        }
    syozoku_map[所属]['count']+= 1
    syozoku_map[所属]['money']+= 金額
print(syozoku_map)


