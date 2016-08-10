import collections
import os
from win32api import GetFileVersionInfo, LOWORD, HIWORD
import xlsxwriter

def get_version_number (filename):
    try:
        info = GetFileVersionInfo (filename, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        return HIWORD (ms), LOWORD (ms), HIWORD (ls), LOWORD (ls)
    except:
        return 0,0,0,0

print("Note: The product path should be under \"C:\Program Files\CyberLink\"")
products = []
while 1:
    product = raw_input("Please enter the product name (case-sensitive) or type 'q' or 'Q' if done: ")
    if product in ['q','Q']:
        break
    else:
        products.append(product)
        
collect = collections.defaultdict(list) #dict type
myPath = "C:\Program Files\CyberLink"
for dirPath, dirNames, fileNames in os.walk(myPath):
    split = dirPath.split('\\') 
    if len(split) >= 4:
        productName = split[3]
        if productName in products:
            for fileName in fileNames:
                fullPath = os.path.join(dirPath, fileName)
                collect[fileName].append(fullPath)
                #print fullPath                
filter = {k: v for k, v in collect.items() if len(v)>1 and (k.split('.')[-1]=="exe" or k.split('.')[-1]=="dll")}

fname = raw_input("Please enter a file name you want to export(Cannot be duplicated!): ")
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(fname+'.xlsx')
worksheet = workbook.add_worksheet()
# Add a format to use to highlight cells.
format = workbook.add_format({'bold': True, 'font_color': 'red'})
# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0
for p in products:
    worksheet.write(0, col, p, format)
    for k, v in filter.items():
        for fp in v:
            if fp.split('\\')[3] == p:
                major,minor,subminor,revision = get_version_number(fp)
                s = "."
                seq = (str(major),str(minor),str(subminor),str(revision))
                ver = s.join(seq)
                #print k, fp, ver
                # Iterate over the data and write it out row by row.
                worksheet.write(row, col,     k)
                worksheet.write(row, col + 1, fp)
                worksheet.write(row, col + 2, ver)
                row += 1
    row = 1
    col += 5

# Write a total using a formula.
#worksheet.write(row, 0, 'Total')
#worksheet.write(row, 1, '=SUM(B1:B4)')
workbook.close()
print("export completed! Go check "+fname+".xlsx")
os.system("pause")

