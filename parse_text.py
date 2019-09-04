#!/usr/bin/python3

import re
import xlsxwriter

workbook = xlsxwriter.Workbook('Bork_Transcript.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Speaker')
worksheet.write('B1', 'Statements')

with open("bork_transcript.txt", "r") as file:
    split = re.split(r"((?:The|Senator|Judge)\s[A-Z]+\W\s)", file.read())
    
    row = 1
    column = 0
    split = split[1:]
    for text in split:
        # Remove extra white space
        text = re.sub(r'\n', ' ', text)
        worksheet.write(row, column, text)
        if column == 1:
            row += 1
        column = (column + 1) % 2

workbook.close()
