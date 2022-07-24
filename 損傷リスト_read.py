import openpyxl
def get_損傷リスト():
    wb = openpyxl.load_workbook('損傷リスト.xlsx')
    ws = wb["Sheet1"]#Sheet1を読み込む


    count_map = {}
    for row in ws.iter_rows(min_row=5):
        所属 = row[3].value 
        if 所属 == None:
            continue
        if 所属 not in count_map:
            count_map[所属] = 0
        count_map[所属] += 1
    #print(count_map)
    return count_map
