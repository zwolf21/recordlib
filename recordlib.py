from collections import OrderedDict, Counter, defaultdict, namedtuple
from heapq import nlargest, nsmallest
from itertools import groupby
from copy import deepcopy
from operator import itemgetter
from io import BytesIO, StringIO
import csv

import xlrd, xlsxwriter



class RecordParser:

	def __init__(self, records=None, columns=None, drop_if=lambda row: False):
		'''dict_list 형태의 데이터셋 전달 drop_if- 제외할 조건전달
			RecordParse(record=[{},{},{}...{}], drop_if=lambda row:bool(row[col])) 
		'''
		if records:
			if not columns:
				fields_set = set()
				for row in records:
					fields_set |= set(row.keys())
			else:
				fields_set = columns
			self.records = [OrderedDict((key, row.get(key, '')) for key in fields_set) for row in records if not drop_if(row)]
		else:
			self.records = []

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


	def to_csv(self, filename=None):
		if not self.records:
			return

		output = StringIO()
		writer = csv.DictWriter(output, fieldnames = self.records[0].keys(), lineterminator='\n')
		writer.writeheader()
		for row in self.records:
			writer.writerow(row)

		if filename:
			with open(filename, 'w') as fp:
				fp.write(output.getvalue())
		else:
			return output.getvalue()

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

	def round_float_fields(self, columns):
		'''
		round_float_fields([('colname1', 2), ('colname2', 3)])
		'''
		for row in self.records:
			for col, rnd in columns:
				try:
					val = float(row[col])
				except:
					continue
				else:
					row[col] = round(val, rnd)
		return self


	def format(self, fmts, drop_if_fail=False):
		'''현재 컬럼의 데이터형을 주어진 기본값과 같은 형으로 설정: fmt=[(대상 컬럼이름, 기본값),()]
			fmt = [('colname1', ''), ('colname2', 0), ('colname3', 0.0)]
		'''
		fmt_funcs = list(map(lambda fmt:(fmt[0] ,type(fmt[1])) , fmts))
		ret = []
		for row in self.records:
			fail = False
			for i, (colname, func) in enumerate(fmt_funcs):
				try:
					val = func(row[colname])
				except:
					fail = True
					val = fmts[i][1]
					row[colname] = val
				else:
					row[colname] = val
			if drop_if_fail and fail:
				continue
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


	def select(self, columns, where=lambda row:True, inplace=True):
		'''select(['A', 'B', 'C', 'D'], where = lambda row: row['A'] > row['B'])
		'''

		if columns == "*" and self.records:
			columns = self.records[0].keys()

		ret = [OrderedDict((key, row[key]) for key in columns) for row in self.records if where(row)]
		if inplace:
			self.records = ret
			return self
		return RecordParser(ret, columns= columns)

	def get_first(self, where, column):
		''' get only one value where = lambda row: row['A'] == 'ABC', column = 'B'
		'''
		for row in self.records:
			if where(row):
				return row[column]


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


	def update(self, bootstrap, where=lambda val:val):
		'''bootstrap = 각 로우의 컬럼을 단계적으로 편집할 (컬럼, 함수)셋, 함수는 해당 로우 값을 인자로 받음 
			bootstrap = [('Columns1', row_func1), ('Columns2', row_func2)...]
		'''
		for row in self.records:
			if not where(row):
				continue
			for column, func in bootstrap:
				row[column] = func(row)
		return self


	def order_by(self, rules):
		'''order_by(['-불출일자', '집계량']), Django Style, 불출일자 내림차순, 집계량 오름차순 
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
			ret.append(OrderedDict((key, row[key]) for key in select))

		if inplace:
			self.records = ret
			return self
		return ret		

	def to2darry(self, headers=True):
		header = [list(self.records[0].keys())]
		body = [list(row.values()) for row in self.records]
		return header + body if headers else body

	def unique(self, column):
		return {row[column] for row in self.records}

	def max(self, column):
		if self.records:
			return max(row[column] for row in self.records)

	def min(self, column):
		if self.records:
			return min(row[column] for row in self.records)

	def value_count(self, column):
		return Counter(row[column] for row in self.records)

	def nlargest_rows(self, num, columns):
		''' 지정된 열을 기준으로 가장 큰 값을 가지는 로우 반환
		recs.nlargest_row(
				num = 3, # 추출할 항목수,
				columns = ['A', 'B'] # 추출 할 열이름
			)
		'''
		return nlargest(num, self.records, key=itemgetter(*columns))

	def nsmallest_rows(self, num, columns):
		''' 지정된 열을 기준으로 가장 작은 값을 가지는 로우 반환
		recs.nlargest_row(
				num = 3, # 추출할 항목수,
				columns = ['A', 'B'] # 추출 할 열이름
			)
		'''
		return nsmallest(num, self.records, key=itemgetter(*columns))

	def get_changes(self, other, pk):
		'''동일한 scheme 를 갖는 RecordParser 객체끼리의 비교
			recs1.get_changes(recs2, pk='primaryKey') 
			pk: 데이터셋 안에서 유일하고 변하지 않는 속성(기본키) 이어야 한다
		'''
		if self.records[0].keys() != other.records[0].keys():
			raise('Other record has different scheme')

		origin, target = {row[pk]: row for row in self}, {row[pk]: row for row in other}

		missing_keys = origin.keys() - target.keys()
		extra_keys = target.keys() - origin.keys()
		common_keys = origin.keys() & target.keys()
		
		added, deleted, updated = [], [], []

		Added = namedtuple('Added', 'index rows')
		Deleted = namedtuple('Deleted', 'index rows')
		Updated = namedtuple('Updated', 'index before after where')

		for key in extra_keys:
			added.append(Added(key, target[key]))

		for key in missing_keys:
			deleted.append(Deleted(key, origin[key]))

		for key in common_keys:
			diff = origin[key].items() ^ target[key].items()
			if diff:
				updated.append(Updated(key, origin[key], target[key], list(dict(diff))))

		Changes = namedtuple('Changes', 'added deleted updated')

		return Changes(added, deleted, updated)

	def _put_changes(self, change_context):
		''' changes = recs.get_changes(recs2, 'pk')
			recs._put_changes(changes)
		'''
		if change_context.added:
			print('---------added---------')
			for added in change_context.added:
				print('	Index:', added.index)

		if change_context.deleted:
			print('--------deleted--------')
			for deleted in change_context.deleted:
				print('	Index:', deleted.index)

		if change_context.updated:
			print('--------updated--------')
			for updated in change_context.updated:
				print('	Index:', updated.index)
				for uk in updated.where:
					print('		where:', uk)
					print('			before:', updated.before[uk])
					print('			after:', updated.after[uk])

	def set_pk(self, columns, pk_name = 'pk'):
		f = lambda row: '-'.join(row[col] for col in columns)
		tmp_recs= RecordParser(self.records)
		tmp_recs.add_column([(pk_name, f)])
		if len(tmp_recs.unique(pk_name)) == len(tmp_recs):
			self.records = tmp_recs.records
			return self
		raise ValueError('The combination of columns {} is duplicated'.format(str(columns)))








def read_excel(file_name=None, file_contents=None, drop_if=lambda row:False, sheet_index=0, start_row=0):
	'''엑셀파일 형태의 데이터 전달 하여 RecordParser 객체 생성 
	'''
	wb = xlrd.open_workbook(filename=file_name, file_contents=file_contents)
	ws = wb.sheet_by_index(sheet_index)
	fields = ws.row_values(start_row)
	records = [OrderedDict(zip(fields, map(str, ws.row_values(r)))) for r in range(start_row+1, ws.nrows)]
	return RecordParser(records, drop_if=drop_if)

def read_csv(filename=None, encoding='utf-8',  fp=None, drop_if=lambda row: False):
	csvfp = None
	if filename:
		csvfp = open(filename, encoding=encoding)
	elif fp:
		csvfp = fp
	else:
		return

	csv_reader = csv.reader(csvfp)
	fields = next(csv_reader)

	records = [OrderedDict(zip(fields, map(str, row))) for row in csv_reader]
	csvfp.close()
	return RecordParser(records, drop_if=drop_if)