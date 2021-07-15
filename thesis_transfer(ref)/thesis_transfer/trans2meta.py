#!/usr/bin/python3
# -*- coding: utf-8 -*-
import pandas as pd
import re, sys, datetime
data_xls = pd.read_excel(sys.argv[1], 'records', index_col=None)
print(len(data_xls))
allcol = list(data_xls)

ncols = ['指導教授姓名(中文)', '中文關鍵詞', '外文關鍵詞']

oldcorr = ['論文名稱(中文)', '論文名稱(外文)', '作者姓名(中文)', '轉檔日期', '類型', '摘要(中文)', '摘要(英文)', '中文關鍵詞', '外文關鍵詞', '語文別', '引用網址', '校院名稱', '系所名稱', '學位類別', '指導教授姓名(中文)', ]

newcorr = ['title', 'title:alternative', 'creator', 'date', 'type', 'description:abstract', '', 'subject1', 'subject2', 'language', 'sys_hyperlink', 'description1', '', 'description2', 'description3']

newcol = ['title', 'title:alternative', 'creator', 'date', 'type', 'description:abstract', 'subject1', 'subject2', 'language', 'sys_hyperlink', 'description1', 'description2', 'description3']



update = {}
mix = [5, 6, 11, 12]

ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')

# 所有需要把換行改成用;或、連接的 data 用的
def NeedTrans(nc, item, name):
	ILLEGAL_CHARACTERS_RE = re.compile(r'[\000-\010]|[\013-\014]|[\016-\037]')	
	# 指導教授需要加上前綴
	if nc == '指導教授姓名(中文)':
		# 如果 cell 有資料，格式會是 str，沒有資料就不是
		if type(item) is str:		
			if '\n' in item:
				# 把 cell data 用換行切開，再用 for 慢慢接成一行
				need = item.split('\n')
				times = 1
				new = "指導教授："
				for word in need:
					new = new + word
					if times < len(need):
						new = new + '、'
					times = times + 1
				
				new = ILLEGAL_CHARACTERS_RE.sub(r'', new)
				update[name].append(new)

			# 沒有換行的就直接加上資料
			else :
				item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
				update[name].append("指導教授：" + item)
		
		# 沒有資料的就不處理，直接加進去
		else :
			item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
			update[name].append(item)

	else:
		if type(item) is str:		
			if '\n' in item:
				need = item.split('\n')
				times = 1
				new = ""
				for word in need:
					new = new + word
					if times < len(need):
						new = new + ';'
					times = times + 1
				new = ILLEGAL_CHARACTERS_RE.sub(r'', new)
				update[name].append(new)
			else :
				item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
				update[name].append(item)
		
		else :
			item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
			update[name].append(item)

# 所有要抓出來的資料
for nc in oldcorr:
	# 需要合併的 column
	if oldcorr.index(nc) in mix:
		if oldcorr.index(nc) == 5:
			# 中英摘要一起抓出來合併
			matchabs = zip(data_xls['摘要(中文)'], data_xls['摘要(英文)'])
			# 用最終的 column name 命名
			new_name = newcorr[oldcorr.index(nc)]
			update[new_name] = []

			for item in matchabs:
				match = str(item[0]) + '\n\n' + str(item[1])
				match = ILLEGAL_CHARACTERS_RE.sub(r'', match)
				update[new_name].append(match)

		elif oldcorr.index(nc) == 11:
			# 校系名稱一起抓出來合併
			matchsch = zip(data_xls['校院名稱'], data_xls['系所名稱'])
			new_name = newcorr[oldcorr.index(nc)]
			update[new_name] = []

			for item in matchsch:
				match = str(item[0]) + str(item[1])
				match = ILLEGAL_CHARACTERS_RE.sub(r'', match)
				update[new_name].append(match)
		else:
			print(":)")
	
	# elif oldcorr.index(nc) == 4:


	# 不需要合併的 column
	else:
		# 用最終的 column name 命名
		new_name = newcorr[oldcorr.index(nc)]
		update[new_name] = []
		
		# 需要重組的 column 
		if nc in ncols:
			for item in data_xls[nc]:
				NeedTrans(nc, item, new_name)

		# 不需要的
		else:
			if nc == '類型':
				for i in range(len(data_xls)):
					update[new_name].append('thesis')

			elif nc == '學位類別':
				for item in data_xls[nc]:
					item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
					update[new_name].append('學位：' + item)
			
			elif nc == '轉檔日期':
				for item in data_xls[nc]:
					if type(item) is str:
						item = ILLEGAL_CHARACTERS_RE.sub(r'', item.split('/')[0])
						update[new_name].append(item)
					else:
						update[new_name].append(item)

			elif nc == '語文別':
				for item in data_xls[nc]:
					if type(item) is str:
						if item == '英文':
							item = 'en'
						else:
							item = 'zh-TW'
						item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
						update[new_name].append(item)
					else:
						update[new_name].append(item)

			else:
				for item in data_xls[nc]:
					item = ILLEGAL_CHARACTERS_RE.sub(r'', item)
					update[new_name].append(item)




df = pd.DataFrame(update)

filename = "transfered.xlsx"

df.to_excel(filename, 'Sheet1', encoding='utf-8-sig', columns=newcol, index=False)

