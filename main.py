from tika import parser
import re
import os
import xlwt

workbook = xlwt.Workbook()

sheet1 = workbook.add_sheet("Version")

sheet1.write(0, 0, "Game ID")
sheet1.write(0, 1, "Version")

row = 1

directory = r'C:\Users\klein\Documents\MyCodes\PyProjects\PdfReader_ADM_Version'
for filename in os.listdir(directory):
    if filename.endswith(".pdf"):
        pdf_Name = filename
        raw_data = parser.from_file(pdf_Name)

        raw_data_trimmed = re.sub('\s+', '', raw_data['content'])

        #print(raw_data_trimmed)

        idHit = raw_data_trimmed.find('eidentificatoconcodiceADM:')
        idEnd = raw_data_trimmed[idHit:].find('Presentata') - 26

        id = raw_data_trimmed[idHit+26:idHit+idEnd+26]

        versionEnd = raw_data_trimmed.find('eidentificatoconcodiceADM:')
        versionStart = raw_data_trimmed.find('Versione:')

        version = raw_data_trimmed[versionStart:versionEnd]

        sheet1.write(row, 0, id)
        sheet1.write(row, 1, version)

        row += 1

workbook.save('Version_List.xls')


