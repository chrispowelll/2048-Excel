# 2048 Excel

import design
import controls
import os
import sys
import random
import openpyxl
import xlwings as xw


def main():
    while True:
        # Original file input
        wbFilename = input("Enter Excel file path: ")

        # Check if user wants to quit
        if wbFilename in ('quit', 'Quit', 'QUIT'):
            sys.exit()

        # Check if file exists
        elif os.path.exists(wbFilename):
            # Create workbook
            design.createValuesWorkbook()
            design.create2048Workbook(wbFilename)
            controls.newGame()

            # Open original workbook and 2048 copy
            updateValues()

            # The Game
            while True:
                move = input()
                if move == 'w':
                    controls.move("up")
                if move == 's':
                    controls.move("down")
                if move == 'a':
                    controls.move("left")
                if move == 'd':
                    controls.move("right")
                if move == 'r':
                    controls.newGame()
                    updateValues()
                if move == 'q':
                    # Open original workbook
                    os.startfile(wbFilename)

                    # Delete game files
                    os.remove("2048.xlsx")
                    os.remove("values.xlsx")

                    # Close program
                    sys.exit()

        # If file doesn't exist
        else:
            print('Invalid file location. Please try again. You can also enter QUIT to exit')


def getNewValue():
    # 80% chance 2 is generated, 20% chance 4 is generated
    numberPicker = random.randint(0, 10)
    if numberPicker >= 8:
        return 4
    else:
        return 2


def updateScore(scoreAdded):
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    ws2048 = wb.active
    ws2048['D1'].value += scoreAdded
    wb.save("values.xlsx")


def updateValues():
    # Open 2048 workbook
    wb2048 = xw.Book("2048.xlsx")
    ws2048 = wb2048.sheets["2048"]

    # Open values workbook
    wbValues = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wbValues.active

    # Update values
    ws2048.range('A2').value = wsValues['A2'].value
    ws2048.range('A3').value = wsValues['A3'].value
    ws2048.range('A4').value = wsValues['A4'].value
    ws2048.range('A5').value = wsValues['A5'].value
    ws2048.range('B2').value = wsValues['B2'].value
    ws2048.range('B3').value = wsValues['B3'].value
    ws2048.range('B4').value = wsValues['B4'].value
    ws2048.range('B5').value = wsValues['B5'].value
    ws2048.range('C2').value = wsValues['C2'].value
    ws2048.range('C3').value = wsValues['C3'].value
    ws2048.range('C4').value = wsValues['C4'].value
    ws2048.range('C5').value = wsValues['C5'].value
    ws2048.range('D1').value = wsValues['D1'].value
    ws2048.range('D2').value = wsValues['D2'].value
    ws2048.range('D3').value = wsValues['D3'].value
    ws2048.range('D4').value = wsValues['D4'].value
    ws2048.range('D5').value = wsValues['D5'].value


if __name__ == "__main__":
    main()
