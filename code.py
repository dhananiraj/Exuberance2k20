import xlrd
import csv

loc = ("R:\\2020\\FEB\\Exuberance\\sheets\\Py\\input_final.xlsx")

wb = xlrd.open_workbook(loc)
for s in range(wb.nsheets):
    sheet = wb.sheet_by_index(s)
    sheetname = sheet.name
    # print(sheetname)
    with open('R:\\2020\\FEB\\Exuberance\\sheets\\{}.csv'.format(sheetname), 'w', newline='') as file:
        writer = csv.writer(file)
        headings = sheet.row_values(0)
        writer.writerow(['Sr. No.']+headings[:7])
        for i in range(1,sheet.nrows):
            r = sheet.row_values(i)
            row1 = [i]+r[:7]
            # print(row1)
            writer.writerow(map(str,row1))
            for j in range(7,64,6):
                bucket = r[j:j+6]
                if bucket == [] :
                    continue
                elif ''.join(map(str,bucket)) == '':
                    continue
                else:
                    # print(bucket)
                    email = bucket.pop(4)
                    bucket.insert(0,email)
                    # bucket.append('')
                    # print(bucket)
                    writer.writerow(map(str,[''] + bucket))
    file.close()
