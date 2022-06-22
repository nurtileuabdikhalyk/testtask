import os
from pathlib import Path
import xlsxwriter

path = os.path.abspath(os.curdir)
def main():
    workbook = xlsxwriter.Workbook('result.xlsx')
    sheet = workbook.add_worksheet()
    sheet.set_column('A:A', 20)
    sheet.set_column('B:B', 40)
    sheet.set_column('C:C', 20)
    sheet.set_column('D:D', 20)
    sheet.write('A1', 'Номер строки')
    sheet.write('B1', 'Папка в которой лежит файл')
    sheet.write('C1', 'название файла')
    sheet.write('D1', 'расширение файла')
    count = 2
    for items in os.walk(path):

        for j in range(len(items[2])):
            sheet.write(f'A{count}', count - 1)
            sheet.write(f'B{count}', items[0])
            sheet.write(f'C{count}', os.path.splitext(items[2][j])[0])
            sheet.write(f'D{count}', Path(items[2][j]).suffix)
            count += 1

    workbook.close()
if __name__ == '__main__':
    main()