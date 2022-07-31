import openpyxl
def get_損傷リスト(file_name):
    wb = openpyxl.load_workbook(file_name)
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
    return syozoku_map
    #print(syozoku_map)
