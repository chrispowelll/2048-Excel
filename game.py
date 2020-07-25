# 2048 Excel

import design
import controls
import os
import sys
import random
import openpyxl


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
            os.startfile("2048.xlsx")

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
                if move == 'q':
                    # Open original workbook
                    os.startfile(wbFilename)

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
    print("update values")


if __name__ == "__main__":
    main()
