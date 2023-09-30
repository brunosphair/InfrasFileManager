import openpyxl
import os

from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule


def get_grd_number(emited_path, ld_name):
    book_path = os.path.join(emited_path, '_LDs', ld_name)
    wb = openpyxl.load_workbook(book_path, read_only=True)
    grd_number = 1
    sheet_name = 'GRD-' + str(grd_number).zfill(3)
    while sheet_name in wb.sheetnames:
        grd_number += 1
        sheet_name = 'GRD-' + str(grd_number).zfill(3)
    wb.close()

    return grd_number


def create_excel_grd(emited_path, ld_name, grd_number, grd_name,
                     ld_information, ld_rev, file_num_caract, grd_items):
    book_path = os.path.join(emited_path, '_LDs', ld_name)
    book = openpyxl.load_workbook(book_path)
    template_sheet = book['GRD-XXX']
    cover_sheet = book['Capa']
    sheet = book.copy_worksheet(template_sheet)
    sheet.title = 'GRD-' + str(grd_number).zfill(3)
    i = 1
    for item in grd_items:
        sheet.cell(row=25 + i, column=1).value = int(i)
        sheet.cell(row=25 + i, column=2).value = item[0]
        sheet.cell(row=25 + i, column=16).value = int(item[1])
        i += 1
    sheet.cell(row=10, column=12).value = grd_name

    sheet.cell(row=7,
                column=12).value = ld_information["emission_date"]

    yellowFill = PatternFill(start_color='FFFF00',
                             end_color='FFFF00',
                             fill_type='solid')
    sheet.conditional_formatting.add('$E$26:$O$192',
                                     FormulaRule(formula=[
                                                 'AND($B26<>"",E26="")'
                                                 ],
                                                 stopIfTrue=False,
                                                 fill=yellowFill
                                                 )
                                    )
    sheet.conditional_formatting.add('$Q$26:$R$192',
                                     FormulaRule(formula=[
                                                 'AND($B26<>"",Q26="")'
                                                 ],
                                                 stopIfTrue=False,
                                                 fill=yellowFill
                                                 )
                                    )

    if ld_rev == -1:
        revision = 0
        ld_name = ld_information["ld_name"]
        cover_sheet.cell(row=2, column=4).value = ld_name
        ld_name = ld_name + "_R0"
        project_title = ld_information["project_title"]
        cover_sheet.cell(row=5, column=1).value = ld_information["ld_title"]
    else:
        revision = ld_rev + 1
        ld_name = ld_name[:file_num_caract] \
            + '_R' \
            + str(revision)
        last_grd = book['GRD-' + str(grd_number - 1).zfill(3)]
        project_title = last_grd.cell(row=1, column=6).value
    sheet.cell(row=1, column=6).value = project_title
    cover_sheet.cell(row=6, column=12).value = revision
    cover_sheet.cell(row=16 + revision, column=1).value = revision
    cover_sheet.cell(row=16 + revision, column=2).value = "C"
    cover_sheet.cell(row=16 + revision, column=3).value = grd_name
    rev_cell = get_cover_cell(revision)
    rev_row = rev_cell[0]
    rev_column = rev_cell[1]

    cover_sheet.cell(row=rev_row,
                     column=rev_column).value = ld_information[
        "emission_date"]
    cover_sheet.cell(row=rev_row + 1,
                     column=rev_column).value = ld_information["acronym1"]
    cover_sheet.cell(row=rev_row + 2,
                     column=rev_column).value = ld_information["acronym2"]
    cover_sheet.cell(row=rev_row + 3,
                     column=rev_column).value = ld_information["acronym3"]

    ld_final_path = os.path.join(emited_path,
                                 '_LDs',
                                 ld_name + '.xlsx')
    book.save(filename=ld_final_path)
    book.close()


def get_cover_cell(rev):
    if rev == 0 or rev == 5:
        column = 3
    elif rev == 1 or rev == 6:
        column = 5
    elif rev == 2 or rev == 7:
        column = 7
    elif rev == 3 or rev == 8:
        column = 8
    elif rev == 4 or rev == 9:
        column = 11

    if rev <= 4:
        row = 32
    else:
        row = 37

    return [row, column]
