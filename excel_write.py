import openpyxl
import excel_read

file_name = '状況表.xlsx'
# ブックを取得
book = openpyxl.load_workbook(file_name)
# シートを取得
sheet = book['4月']
# セルへ書き込む
sheet['D19'] = 'ナンバー'



# 保存する
book.save(file_name)


