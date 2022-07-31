import openpyxl
def find_row_number(ws,column_number,key):
    row_number = 0
    
    for row in ws.iter_rows():
        row_number +=1
        所属 = row[column_number].value
        if 所属 == None:
            continue
        所属 = 所属.replace('　', '')
        print(row_number,所属)
        if key == 所属:
            return row_number 
    
    
def write_状況表(損傷リスト):
    file_name = '状況表.xlsx'
    # ブックを取得
    book = openpyxl.load_workbook(file_name)
    # シートを取得
    sheet = book['４月']
    # セルへ書き込む
    #sheet['D19'] = 損傷リスト['浦和']
    for key in 損傷リスト:
        row_number = find_row_number(sheet,1,key) #1はB列のこと
        if row_number == None:
            continue
        cell_number = 'D'+str(row_number)
        sheet[cell_number] = 損傷リスト[key]['count']
        cell_number = 'G'+str(row_number)
        sheet[cell_number] = 損傷リスト[key]['money']
        #print(row_number)
    for key in 損傷リスト:
        row_number = find_row_number(sheet,23,key) #23はX列のこと
        if row_number == None:
            continue
        cell_number = 'Z'+str(row_number)
        sheet[cell_number] = 損傷リスト[key]['count']
        cell_number = 'AC'+str(row_number)
        sheet[cell_number] = 損傷リスト[key]['money']
        
    # 保存する
    book.save(file_name)
