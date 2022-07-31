import 損傷リスト_read
import 状況表_write

損傷リスト = 損傷リスト_read.get_損傷リスト('損傷リスト.xlsx')
print(損傷リスト)
状況表_write.write_状況表(損傷リスト,'状況表.xlsx','５月')

