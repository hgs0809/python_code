import xlrd,xlwt

data = xlrd.open_workbook('1.xlsx')
workboot = xlwt.Workbook(encoding='ascii')
worksheet = workboot.add_sheet("1")
all_table = {}
sheet_l = data.sheets()
n_row = len(sheet_l)
for n in range(0,n_row):
	sheet = sheet_l[n]
	all_table["table_%d"%n] = {}
	for z in range(1,sheet.nrows):
		all_table["table_%d"%n][sheet.cell(z,1).value] = sheet.cell(z,8).value

new_table = {}
for v in all_table.values():
	for kv,vv in v.items():
		if new_table.has_key(kv):
			if new_table[kv] < vv:
				new_table[kv] = vv
		else:
			new_table[kv] = vv

num = 0
for k,v in new_table.items():
	print k,v
	worksheet.write(num,0,k)	
	worksheet.write(num,1,v)	
	num +=1
workboot.save('b.xlsx')
