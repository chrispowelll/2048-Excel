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

                    # Close program
                    sys.exit()

        # If file doesn't exist
        else:
            print('Invalid file location. Please try again. You can also enter QUIT to exit')


def getNewValue():
    # 90% chance 2 is generated, 10% chance 4 is generated
    numberPicker = random.randint(0, 10)
    if numberPicker >= 9:
        return 4
    else:
        return 2


def updateScore(scoreAdded):
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    ws2048 = wb.active
    ws2048['D1'].value += scoreAdded
    wb.save("values.xlsx")


def updateValues():
    gameCells = ['A2', 'A3', 'A4', 'A5', 'B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5']

    # Open 2048 workbook
    wb2048 = xw.Book("2048.xlsx")
    ws2048 = wb2048.sheets["2048"]

    # Open values workbook
    wbValues = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wbValues.active

    # Update score
    ws2048.range('D1').value = wsValues['D1'].value

    # Update game status
    ws2048.range('A6').value = wsValues['A6'].value

    # Update values
    for i in gameCells:
        # Only show values greater than 0
        if wsValues[i].value > 0:
            ws2048.range(i).value = wsValues[i].value
        else:
            ws2048.range(i).value = ""


if __name__ == "__main__":
    main()
