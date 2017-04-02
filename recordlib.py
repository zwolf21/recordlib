from collections import OrderedDict
from itertools import groupby
from copy import deepcopy
from operator import itemgetter
from io import BytesIO

import xlrd, xlsxwriter


def read_excel(file_name=None, file_contents=None, drop_if=lambda row:False, sheet_index=0, start_row=0):
	'''엑셀파일 형태의 데이터 전달 하여 RecordParser 객체 생성 
	'''
	wb = xlrd.open_workbook(filename=file_name, file_contents=file_contents)
	ws = wb.sheet_by_index(sheet_index)
	fields = ws.row_values(start_row)
	records = [OrderedDict(zip(fields, map(str, ws.row_values(r)))) for r in range(start_row+1, ws.nrows)]
	return RecordParser(records, drop_if)


class RecordParser:
	def __init__(self, records=None, drop_if=lambda row: False):
		'''dict_list 형태의 데이터셋 전달 drop_if- 제외할 조건전달
			RecordParse(record=[{},{},{}...{}], drop_if=lambda row:bool(row[col])) 
		'''
		if records:
			self.records = [row for row in records if not drop_if(row)]

	def __getitem__(self, index):
		if self.records:
			return self.records[index]

	def __len__(self):
		return len(self.records)

	def __add__(self, other):
		return RecordParser(self.records+other.records)

	def __iadd__(self, other):
		self.records += other.records
		return self

	def __iter__(self):
		return iter(self.records)


	def read_excel(self, file_name=None, file_contents=None, drop_if=lambda row:False, sheet_index=0, start_row=0):
		'''엑셀파일 형태의 데이터 전달 
		'''
		wb = xlrd.open_workbook(filename=file_name, file_contents=file_contents)
		ws = wb.sheet_by_index(sheet_index)
		fields = ws.row_values(start_row)
		records = [OrderedDict(zip(fields, map(str, ws.row_values(r)))) for r in range(start_row+1, ws.nrows)]
		self.records = [row for row in records if not drop_if(row)]
		return self

	def to_excel(self, filename=None):
		'''filename 을 전달하지 않으면 file contents 를 반환
		'''
		if not self.records:
			return

		output = BytesIO()
		wb = xlsxwriter.Workbook(output)
		ws = wb.add_worksheet()
		ws.write_row(0,0, self.records[0].keys())
		for r, row in enumerate(self.records, 1):
			ws.write_row(r,0, row.values())
		wb.close()
		if filename:
			with open(filename, 'wb') as fp:
				fp.write(output.getvalue())
		else:
			return output.getvalue()


	def format(self, fmts, drop_if_fail=False):
		'''현재 컬럼의 데이터형을 주어진 기본값과 같은 형으로 설정: fmt=[(대상 컬럼이름, 기본값),()]
			fmt = [('colname1', ''), ('colname2', 0), ('colname3', 0.0)]
		'''
		fmt_funcs = list(map(lambda fmt:(fmt[0] ,type(fmt[1])) , fmts))
		ret = []
		for row in self.records:
			for i, (colname, func) in enumerate(fmt_funcs):
				try:
					val = func(row[colname])
				except:
					val = fmts[i][1]
				else:
					pass
				finally:
					row[colname] = val
					if not drop_if_fail:
						ret.append(row)
		self.records = ret
		return self


	def rename(self, renames):
		'''컬럼의 이름을 변경
			renames = [('OriColname', 'newColname'), ('OriColname2', 'newColname2')]
		'''
		columns = list(self.records[0].keys())
		for i, col in enumerate(columns):
			for old, new in renames:
				if old == col:
					columns[i] = new

		for row in self.records:
			for old, new in renames:
				row[new] = row.pop(old)

		self.select(columns)
		return self

	def vlookup(self, foreign, fk, pk, ret_columns):
		'''새로운 컬럼을 지정하여 다른 테이블에서 정보를 찾아서 채움
			fk = 현재 테이블에서 참조할 컬럼, pk = 다른테이블의 컬럼, ret_column =다른테이블의 가져올 정보가 있는 컬럼 및 디폴트 값 지정
			vlookup(RecordParser2, 'FK', 'PK', [('RetColumn1', 0),('RetColumn2', "")])
		'''
		if not foreign:
			return self

		foreign = {row[pk] :row for row in foreign}

		for self_row in self.records:
			foreign_row = foreign.get(self_row[fk], {})
			for col, default in ret_columns:
				self_row[col] = foreign_row.get(col, default)
		return self


	def value_map(self, mappings):
		'''특정 값을 정해진 값으로 변경 
			mappings = [('Column1', {'1' :'일반', '2': '프리미엄', '3': 'VIP'}, ETC1), ('Columns2', {'A':'특급', 'B':'중급', 'C':하급}, ETC2)]
			ETC : 매핑에 값이 존재 하지 않을경우 기본값
		'''
		for row in self.records:
			for column, value_map, etc in mappings:
				row[column] = value_map.get(row[column])
		return self


	def select(self, columns, where=lambda row:True):
		'''select([('A', 'B', 'C', 'D')], where = lambda row: row['A'] > row['B'])
		'''
		if columns == "*":
			columns = self.records[0].keys()
		self.records = [OrderedDict((key, row[key]) for key in columns) for row in self.records if where(row)]
		return self


	def add_column(self, columns=[]):
		'''columns = [('colname1', row_func1), ('colname2', row_func2)....]
			add_column([('불출월', lambda row:row['불출일자'][:-3])])
		'''
		for row in self.records:
			for colname, func in columns:
				row[colname] = func(row)
		return self


	def drop_column(self, columns=[]):
		for row in self.records:
			for column in columns:
				row.pop(column)
		return self


	def update(self, bootstrap, where=lambda row:row):
		'''bootstrap = 각 로우의 컬럼을 단계적으로 편집할 (컬럼, 함수)셋, 함수는 로우 값을 인자로 받음 
			bootstrap = [('Columns1', row_func1), ('Columns2', row_func2)...]
		'''
		for row in self.records:
			if not where(row):
				continue
			for column, func in bootstrap:
				row[column] = func(row)
		return self


	def order_by(self, rules):
		'''order_by('-불출일자', '집계량'), Django Style, 불출일자 내림차순, 집계량 오름차순 
		'''
		for rule in reversed(rules):
			rvs = rule.startswith('-')
			rule = rule.strip('-')
			self.records.sort(key=lambda x: x[rule], reverse=rvs)
		return self


	def distinct(self, columns, eliminate=False):
		'''columns 에서 지정한 컬럼을 기준으로 유일값만 추출,
			eliminate = 중복항목 전부 제거, 중복된 경우가 없는 로우만 남는다.
		'''
		ret = []
		for g, l in groupby(sorted(self.records, key=itemgetter(*columns)), key=itemgetter(*columns)):
			head, *body = l
			if eliminate and body:
				continue
			else:
				ret.append(head)
		self.records = ret
		return self


	def group_by(self, columns, aggset, selects=[], inplace=True):
		'''ret = recs.group_by(
				columns = ['불출일자', '병동'], 
				aggset = [('집계량', sum, '집계량종합'), ('집계량', len, '개수')], 
				selects=['불출일자', '약품코드', '집계량종합', '개수']
			)
			columns : 그룹핑 기준 컬럼들 (다중그루핑 가능),
			aggset : 집계를 구할 컬럼, 집계함수, 집계 결과값을 받을 컬럼 이름,
			selects : 지정 안할 시 기본적인 집계 결과, 지정시 selects 컬럼에 해당하는 값 중 각 집계된 항목들 중에서 첫번째 값을 가져옴
		'''
		ret = []
		for g, l in groupby(sorted(self.records, key=itemgetter(*columns)), key=itemgetter(*columns)):
			grouped = list(l)
			row = grouped[0]		
			for aggcol, aggfunc, alias in aggset:
				row[alias] = aggfunc([row[aggcol] for row in grouped])

			select = selects if selects else columns+[e[2] for e in aggset]
			ret.append(OrderedDict((key, val) for key, val in row.items() if key in select))

		if inplace:
			self.records = ret
			return self
		return ret

	def to2darry(self, header=True):
		header = [list(self.records[0].keys())]
		body = [list(row.values()) for row in self.records]
		return header + body if header else body



