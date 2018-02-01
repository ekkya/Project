import xlsxwriter
from xlrd import open_workbook
from scipy import signal

row = 0
col = 0
list = []
wb = open_workbook('data_PC_NBC_UTPC_UTPC.xlsx')
sheet = wb.sheet_by_index(0)
workbook = xlsxwriter.Workbook('data_PC_NBC_UTPC_UTPC_u.xlsx')
worksheet = workbook.add_worksheet()
i = 0
for row_num in range(sheet.nrows):
    row_value = sheet.row_values(row_num)
    if row_value[1] > "0":
        #print row_value
        s = [str(item) for item in row_value]
        print s
        list.append(s[0])
        i = i + 1
        worksheet.write(row, col, s[0])
        worksheet.write(row, col + 1, s[1])
        worksheet.write(row, col + 2, 'PC_NBC_UTPC_UTPC')
        worksheet.write(row, col + 3, 1 / (float(s[0]) * 0.001))
        row += 1
    #row = 1
    worksheet.write(row, col + 4, i)
#print len(list)
results = map(float, list)
#print len(results)
r,fx = signal.periodogram(results)
i = sum(fx)/len(fx)
row = 0
col = 0
worksheet.write(row, col + 7, i)
print "Complete"
workbook.close()

row = 0
col = 0
list = []
wb = open_workbook('data_PC_NBC_UTPC_UTPC_2.xlsx')
sheet = wb.sheet_by_index(0)
workbook = xlsxwriter.Workbook('data_PC_NBC_UTPC_UTPC_2_u.xlsx')
worksheet = workbook.add_worksheet()
i = 0
for row_num in range(sheet.nrows):
    row_value = sheet.row_values(row_num)
    if row_value[1] > "0":
        #print row_value
        s = [str(item) for item in row_value]
        print s
        list.append(s[0])
        i = i + 1
        worksheet.write(row, col, s[0])
        worksheet.write(row, col + 1, s[1])
        worksheet.write(row, col + 2, 'PC_NBC_UTPC_UTPC')
        worksheet.write(row, col + 3, 1 / (float(s[0]) * 0.001))
        row += 1
    #row = 1
    worksheet.write(row, col + 4, i)
#print len(list)
results = map(float, list)
#print len(results)
r,fx = signal.periodogram(results)
i = sum(fx)/len(fx)
row = 0
col = 0
worksheet.write(row, col + 7, i)
print "Complete"
workbook.close()

row = 0
col = 0
list = []
wb = open_workbook('data_PC_NBC_UTPC_UTPC_3.xlsx')
sheet = wb.sheet_by_index(0)
workbook = xlsxwriter.Workbook('data_PC_NBC_UTPC_UTPC_3_u.xlsx')
worksheet = workbook.add_worksheet()
i = 0
for row_num in range(sheet.nrows):
    row_value = sheet.row_values(row_num)
    if row_value[1] > "0":
        #print row_value
        s = [str(item) for item in row_value]
        print s
        list.append(s[0])
        i = i + 1
        worksheet.write(row, col, s[0])
        worksheet.write(row, col + 1, s[1])
        worksheet.write(row, col + 2, 'PC_NBC_UTPC_UTPC')
        worksheet.write(row, col + 3, 1 / (float(s[0]) * 0.001))
        row += 1
    #row = 1
    worksheet.write(row, col + 4, i)
#print len(list)
results = map(float, list)
#print len(results)
r,fx = signal.periodogram(results)
i = sum(fx)/len(fx)
row = 0
col = 0
worksheet.write(row, col + 7, i)
print "Complete"
workbook.close()
