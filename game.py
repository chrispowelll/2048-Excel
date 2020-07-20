# 2048 Excel

import design
import controls
import os
import sys
import random
import openpyxl


def main():
    # Original file input
    wbFilename = input("Enter Excel file path: ")

    # Create workbook
    design.createWorkbook(wbFilename)
    controls.startGame()

    # Open original workbook and 2048 copy
    os.startfile("2048.xlsx")

    # The Game
    while True:
        move = input()
        if move == 'w':
            controls.moveUp()
        if move == 's':
            controls.moveDown()
        if move == 'a':
            controls.moveLeft()
        if move == 'd':
            controls.moveRight()
        if move == 'r':
            controls.startGame()
        if move == 'q':
            os.startfile(wbFilename)
            sys.exit()
        if move in ['r', 'd', 'a', 's', 'w']:
            updateScore()
            print("update values")


def getNewValue():
    numberPicker = random.randint(0, 10)
    if numberPicker >= 8:
        return 4
    else:
        return 2


def updateScore():
    score = 0
    wb = openpyxl.load_workbook("2048.xlsx")
    ws2048 = wb.active
    ws2048['D1'] = str(score)


if __name__ == "__main__":
    main()
