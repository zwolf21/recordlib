from recordlib import RecordParser, read_excel


path1 = r'C:\Users\HS\Desktop\향정잔량.xls'
path2 = "C:\\Users\\HS\\Desktop\\향정잔량2.xls"
path3 = "C:\\Users\\HS\\Desktop\\약품정보.xls"
path4 = "C:\\Users\\HS\\Desktop\\약품정보2.xls"



# recs = read_excel(path1, drop_if=lambda row: row['불출일자'] == '')
# recs2 = read_excel(path2, drop_if=lambda row: row['불출일자'] == '')
recs3 = read_excel(path3)
recs4 = read_excel(path4)
# recs.round_float_fields([('총량', 0), ('처방량(규격단위)', 2), ('불출량', 2)])

# for row in recs[:1]:
# 	print(row['처방량(규격단위)'])

# # print(recs.nsmallest_rows(3, ['병동', '총량']))
info =  recs3.get_changes(recs4, '약품코드')
print(info.updated[0].where)
# info = recs.get_changes(recs2, ['불출일자', '병동', '약품코드'])
# recs3._put_changes(info)
# recs3.set_pk(['투여경로', '제약회사명'])
# recs.set_pk(['접수일련번호', '접수처방일련번호'])
# for row in recs:
# 	print(row['pk'])
