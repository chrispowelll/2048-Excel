# Creates 2048 worksheet and design

import openpyxl
from openpyxl.styles import Font


def createWorkbook(wbFilename):
    # Open Excel workbook
    wb = openpyxl.load_workbook(wbFilename)

    # Create 2048 workbook copy
    wb.save("2048.xlsx")

    # Create 2048 worksheet
    ws2048 = wb.copy_worksheet(wb.active)
    ws2048.title = "2048"

    # Delete other sheets
    sheets = wb.get_sheet_names()
    sheetName = wb.sheetnames
    for i in range(0, len(sheets)):
        if sheets[i] != "2048":
            wb.remove(wb[sheetName[i]])

    # Set fonts
    for i in ['A1', 'B1', 'C1', 'D1', 'E1']:  # headings
        ws2048[i].font = Font(name='Verdana', size=10, bold=True)
    for i in ['A2', 'A3', 'A4', 'A5', 'B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5']:  # game
        ws2048[i].font = Font(name='Verdana', size=10)
    for i in ['E2', 'E3', 'E4', 'E5']:  # controls
        ws2048[i].font = Font(name='Verdana', size=8)

    # Input headings and controls
    ws2048['A1'] = "2048"
    ws2048['C1'] = "Score:"
    ws2048['E1'] = "Controls"
    ws2048['E2'] = "W = Up, S = Down"
    ws2048['E3'] = "A = Left, D = Right"
    ws2048['E4'] = "R = Reset"
    ws2048['E5'] = "Q = Quit"
    blankCells = ['B1', 'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'A6', 'B6', 'C6', 'D6', 'E6', 'G6', 'G5', 'G4', 'G3', 'G2', 'G1']
    for i in range(0, len(blankCells)):
        ws2048[blankCells[i]] = ""

    # Save workbook
    wb.save("2048.xlsx")
