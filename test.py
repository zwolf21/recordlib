from recordlib import RecordParser, read_excel


path1 = r'C:\Users\HS\Desktop\향정.xls'
path2 = "C:\\Users\\HS\\Desktop\\미불출.xls"
path3 = "C:\\Users\\HS\\Desktop\\약품정보.xls"



recs = read_excel(path1)
# recs2 = read_excel(path2)
di_recs = read_excel(path3)
# di_recs.format([('함량1', 0.0), ('일반단가' ,0)])

# recs+=recs2
# recs.read_excel(path2)
# recs.format([('약국진행상태', 0.0), ('처방량(규격단위)', 0.0), ('집계량', 0.0)])
recs.select(['불출일자', '원처방일자', '처방번호[묶음]', '병동','환자명','환자번호','약품명', '약품코드', '처방량(규격단위)', '집계량'], 
	# lambda row: '아네폴' in row['약품명'] or '듀로제' in row['약품명']
	# lambda row: '아네폴' in row['약품명']
	# lambda row: '듀로제' in row['약품명']
	# lambda row: '[퇴원]' in row['약품명']
	# lambda row: row['환자명'] == '박영자'
)
# recs.rename([('원처방일자', '원숭일자')])
recs.vlookup(di_recs, '약품코드' ,'약품코드', [('일반단가', 0), ('투여경로', ""), ('함량1', 0), ('함량단위1', "")])
# recs.value_map([('투여경로', {'1': '내복약', '2': '외용약', '3':'주사약'}, '구분없음')])

recs.update(bootstrap = [
	('불출일자', lambda row: row['불출일자'][:-3])
])

# recs.add_column([('불출월', lambda row:row['불출일자'][:-3])])

# recs.order_by(['-불출일자', '집계량'])

# # recs.distinct(['불출일자','환자번호','약품명','처방번호[묶음]'], True)

# ret = recs.group_by(
# 	columns = ['불출일자', '병동', '약품명'], 
# 	aggset = [('집계량', sum, '집계량종합'), ('집계량', len, '개수')], 
# 	selects=['불출일자', '병동', '약품명', '집계량종합', '개수']
# )

ret = recs.to2darry()

print(ret[:2])
# for rec in recs.records:
# 	print(rec)

# print(len(recs))

# recs.to_excel('test.xlsx')