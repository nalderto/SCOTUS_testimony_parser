#!/usr/bin/python3

import re
import xlsxwriter

def addTranscript(worksheet, fileName, row, order):
    column = 0
    with open(fileName, "r") as file:
        split = re.split(r"([A-Z]+:)", file.read())
        split = split[1:]
        for text in split:
            # Remove extra white space
            text = re.sub(r'\n', ' ', text)
            text = re.sub(r'\s\s+', ' ', text)
            worksheet.write(row, column, text)
            if column == 1:
                worksheet.write(row, 2, order)
                worksheet.write(row, 3, fileName.split(".")[0])
                row += 1
                order += 1
            column = (column + 1) % 2
    return row, order

def main():

    workbook = xlsxwriter.Workbook('Kavanaugh_Transcript.xlsx')
    worksheet = workbook.add_worksheet()


    worksheet.write('A1', 'Speaker')
    worksheet.write('B1', 'Statements')
    worksheet.write('C1', 'Order')
    worksheet.write('D1', 'Day')

    row = 1
    order = 1

    row, order = addTranscript(worksheet, "Day_1.txt", row, order)
    row, order = addTranscript(worksheet, "Day_2_Part_1.txt", row, order)
    row, order = addTranscript(worksheet, "Day_2_Part_1.txt", row, order)
    row, order = addTranscript(worksheet, "Day_3_Part_1.txt", row, order)
    row, order = addTranscript(worksheet, "Day_3_Part_2.txt", row, order)
    row, order = addTranscript(worksheet, "Day_3_Part_2.txt", row, order)
    row, order = addTranscript(worksheet, "Day_5_Part_2.txt", row, order)
    workbook.close()


if __name__ == "__main__":
    main()