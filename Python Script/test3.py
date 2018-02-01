import xlsxwriter
import json

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('data_L6_prop.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

# Some data we want to write to the worksheet.
data = json.load(open('layer_download.json'))
#print data
l6_a = data["L6"]["total axonal length"]
print l6_a
l6_d = data["L6"]["total dendritic length"]
print l6_d
l6_ep =  data["L6"]["number of excitatory pathways"]
print l6_ep
l6_ip =  data["L6"]["number of interlaminar pathways"]
print l6_ip

worksheet.write(row, col, l6_a)
worksheet.write(row, col + 1, l6_d)
worksheet.write(row, col + 2, l6_ep)
worksheet.write(row, col + 3, l6_ip)

print "Complete"
workbook.close()

workbook = xlsxwriter.Workbook('data_L23_prop.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

# Some data we want to write to the worksheet.
data = json.load(open('layer_download.json'))
#print data
l23_a = data["L23"]["total axonal length"]
print l23_a
l23_d = data["L23"]["total dendritic length"]
print l6_d
l23_ep =  data["L23"]["number of excitatory pathways"]
print l23_ep
l23_ip =  data["L23"]["number of interlaminar pathways"]
print l23_ip

worksheet.write(row, col, l23_a)
worksheet.write(row, col + 1, l23_d)
worksheet.write(row, col + 2, l23_ep)
worksheet.write(row, col + 3, l23_ip)

print "Complete"
workbook.close()

workbook = xlsxwriter.Workbook('data_L5_prop.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

# Some data we want to write to the worksheet.
data = json.load(open('layer_download.json'))
#print data
l5_a = data["L5"]["total axonal length"]
print l5_a
l5_d = data["L5"]["total dendritic length"]
print l5_d
l5_ep =  data["L5"]["number of excitatory pathways"]
print l5_ep
l5_ip =  data["L5"]["number of interlaminar pathways"]
print l5_ip

worksheet.write(row, col, l5_a)
worksheet.write(row, col + 1, l5_d)
worksheet.write(row, col + 2, l5_ep)
worksheet.write(row, col + 3, l5_ip)

print "Complete"
workbook.close()

workbook = xlsxwriter.Workbook('data_L4_prop.xlsx')
worksheet = workbook.add_worksheet()

row = 0
col = 0

# Some data we want to write to the worksheet.
data = json.load(open('layer_download.json'))
#print data
l4_a = data["L4"]["total axonal length"]
print l4_a
l4_d = data["L4"]["total dendritic length"]
print l4_d
l4_ep =  data["L4"]["number of excitatory pathways"]
print l4_ep
l4_ip =  data["L4"]["number of interlaminar pathways"]
print l4_ip

worksheet.write(row, col, l4_a)
worksheet.write(row, col + 1, l4_d)
worksheet.write(row, col + 2, l4_ep)
worksheet.write(row, col + 3, l4_ip)

print "Complete"
workbook.close()