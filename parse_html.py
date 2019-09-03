import lxml.html as lh
import xlsxwriter
import re

html = open("Bork_Table.html", "r")

doc = lh.fromstring(html.read())

speakers = doc.xpath('//td/strong')
statements = doc.xpath('//td/p')

workbook = xlsxwriter.Workbook('Bork_Table.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Speaker')
worksheet.write('B1', 'Statements')

row = 1
for speaker in speakers:
    worksheet.write(row, 0, speaker.text_content())
    worksheet.write(row, 1, re.sub(r'\s\s+', ' ', statements[row - 1].text_content()))
    row += 1

workbook.close()