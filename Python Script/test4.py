import xlsxwriter

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('data_PC_PC_STPC_IPC.xlsx')
worksheet = workbook.add_worksheet()

# Some data we want to write to the worksheet.
row = 0
col = 0

with open('PC_PC_STPC_IPC.dat', 'r') as f:
    data = f.readlines()
    #print data
    for line in data:
        words = line.split()
        worksheet.write(row, col, words[0])
        worksheet.write(row, col + 1, words[1])
        row += 1

print "Complete"
workbook.close()

workbook = xlsxwriter.Workbook('data_PC_PC_STPC_IPC_2.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

with open('PC_PC_STPC_IPC_2.dat', 'r') as f:
    data = f.readlines()
    #print data
    for line in data:
        words = line.split()
        worksheet.write(row, col, words[0])
        worksheet.write(row, col + 1, words[1])
        row += 1

print "Complete"
workbook.close()

workbook = xlsxwriter.Workbook('data_PC_PC_STPC_IPC_3.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

with open('PC_PC_STPC_IPC_3.dat', 'r') as f:
    data = f.readlines()
    #print data
    for line in data:
        words = line.split()
        worksheet.write(row, col, words[0])
        worksheet.write(row, col + 1, words[1])
        row += 1

print "Complete"
workbook.close()