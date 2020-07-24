import openpyxl
import game
import random


def newGame():
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wb.active
    gameCells = ['A2', 'A3', 'A4', 'A5', 'B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5']

    # Fill board with 0s
    for i in gameCells:  # game
        wsValues[i] = 0

    # Reset score
    wsValues['D1'] = 0
    wsValues['A2'] = 2
    wsValues['A3'] = 4
    wsValues['A4'] = 4
    wsValues['A5'] = 2

    # Create random starting block
    randomCell = random.choice(gameCells)
    wsValues[randomCell] = str(game.getNewValue())

    # Save and update values
    wb.save("values.xlsx")
    game.updateValues()


def move(direction):
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wb.active
    if direction == "up":
        if wsValues['A2'].value == wsValues['A3'].value:
            wsValues['A2'] = wsValues['A3'].value * 2
            wsValues['A3'] = wsValues['A4'].value
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A3'].value == wsValues['A4'].value:
            wsValues['A3'] = wsValues['A4'].value * 2
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A4'].value == wsValues['A5'].value:
            wsValues['A4'] = wsValues['A5'].value * 2
            wsValues['A5'] = 0
        if wsValues['B2'].value == wsValues['B3'].value:
            wsValues['B2'] = wsValues['B3'].value * 2
            wsValues['B3'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B3'].value == wsValues['B4'].value:
            wsValues['B3'] = wsValues['B4'].value * 2
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B4'].value == wsValues['B5'].value:
            wsValues['B4'] = wsValues['B5'].value * 2
            wsValues['B5'] = 0
        if wsValues['C2'].value == wsValues['C3'].value:
            wsValues['C2'] = wsValues['C3'].value * 2
            wsValues['C3'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C3'].value == wsValues['C4'].value:
            wsValues['C3'] = wsValues['C4'].value * 2
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C4'].value == wsValues['C5'].value:
            wsValues['C4'] = wsValues['C5'].value * 2
            wsValues['C5'] = 0
        if wsValues['D2'].value == wsValues['D3'].value:
            wsValues['D2'] = wsValues['D3'].value * 2
            wsValues['D3'] = wsValues['D4'].value
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D3'].value == wsValues['D4'].value:
            wsValues['D3'] = wsValues['D4'].value * 2
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D4'].value == wsValues['D5'].value:
            wsValues['D4'] = wsValues['D5'].value * 2
            wsValues['D5'] = 0
    elif direction == "down":
        print("move down")
    elif direction == "left":
        print("move left")
    elif direction == "right":
        print("move right")

    # Save and update values
    wb.save("values.xlsx")
    game.updateValues()
