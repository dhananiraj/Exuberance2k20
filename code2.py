import xlrd
import csv

loc = ("R:\\2020\\FEB\\Exuberance\\sheets\\Py\\raw_file.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
GroupEvents = ['Story Writing (3)', 'Story Writing (Eng.)(compulsory 3)', 'Story Writing(Guj.) (Compulsory 3)', 'Ad Mad Show (5)', 'Selfie Props (4)', 'Collage (4)', 'Stone Painting (2)', 'Rangoli (4)']
SoloEvents = ['Debate (Gujarati)', 'Debate (English)', 'Elocution (Hindi)', 'Cast a Spell', 'Elocution (English)', 'Mehendi', 'Rangoli', 'Poem (Gujarati)', 'Poem (Hindi)', 'Debate (Hindi)', 'Elocution (Gujarati)', 'Poster Competition', 'Poem (English)']

sheet = wb.sheet_by_index(0)
for event in GroupEvents:
    event = event.strip()
    with open('R:\\2020\\FEB\\Exuberance\\sheets\\{}.csv'.format(event), 'w', newline='') as file:
        writer = csv.writer(file)
        headings = sheet.row_values(0)
        headings.pop(0)
        headings.pop(6)
        headings.pop(6)
        writer.writerow(['Sr. No.']+headings[:7]+['Sign'])
        sr = 1
        for i in range(1,sheet.nrows):
            r = sheet.row_values(i)
            if r[8].strip() == event:
                r.pop(0)
                r.pop(7-1)
                r.pop(8-2)
                row1 = [sr]+r[:7]
                # print(row1)
                writer.writerow(map(str,row1))
                for j in range(7,70,6):
                    bucket = r[j:j+6]
                    if bucket == [] :
                        continue
                    elif ''.join(map(str,bucket)) == '':
                        continue
                    else:
                        email = bucket.pop(4)
                        bucket.insert(0,email)
                        writer.writerow(map(str,[''] + bucket))
                sr+=1
    file.close()

for event in SoloEvents:
    event = event.strip()
    with open('R:\\2020\\FEB\\Exuberance\\sheets\\{}.csv'.format(event), 'w', newline='') as file:
        writer = csv.writer(file)
        headings = sheet.row_values(0)
        headings.pop(0)
        headings.pop(6)
        headings.pop(6)
        writer.writerow(['Sr. No.']+headings[:6]+['Sign'])
        sr = 1
        for i in range(1,sheet.nrows):
            r = sheet.row_values(i)
            if r[-1].find(event) >= 0:
                r.pop(0)
                r.pop(7-1)
                r.pop(8-2)
                row1 = [sr]+r[:7]
                # print(row1)
                writer.writerow(map(str,row1))
                sr+=1
    file.close()
