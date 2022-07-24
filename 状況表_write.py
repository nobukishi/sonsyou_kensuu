import openpyxl

def write_状況表(損傷リスト):
    file_name = '状況表.xlsx'
    # ブックを取得
    book = openpyxl.load_workbook(file_name)
    # シートを取得
    sheet = book['４月']
    # セルへ書き込む
    sheet['D19'] = 損傷リスト['浦和']



    # 保存する
    book.save(file_name)
