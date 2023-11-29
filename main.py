import openpyxl
from win32com import client

book = openpyxl.open('data_owners.xlsx')
sheet = book.active
new_row = 1


for row in range(10, 15):
    new_book = openpyxl.load_workbook('reshenie_1.xlsx')
    new_sheet = new_book.active

    number_room = str(sheet[row][0].value)
    fio = str(sheet[row][1].value)
    room_area = float(sheet[row][2].value)
    part_own_str = str(sheet[row][4].value)
    part_own = sheet[row][4].value
    if type(part_own) == str:
        a, b = part_own.split("/")
        part_own = int(a) / int(b)
    document_number = str(sheet[row][5].value)
    new_sheet['C6'].value = number_room
    new_sheet['A4'].value = fio
    new_sheet['C9'].value = room_area
    new_sheet['C10'].value = part_own_str
    new_sheet['C11'].value = part_own * room_area
    new_sheet['A7'].value = document_number

    
    new_book.save(f'Protocol{new_row}.xlsx')
    new_book.close()
    
    excel = client.Dispatch("Excel.Application")

    sheets = excel.Workbooks.Open(f'C:\\my_program\\Protocol{new_row}.xlsx')

    work_sheets = sheets.Worksheets[0]

    work_sheets.ExportAsFixedFormat(0, f"C:\\my_program\\pdf\\mypdf{new_row}.pdf")
    new_row += 1