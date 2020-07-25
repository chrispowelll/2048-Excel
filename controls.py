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

    # Create random starting block
    randomCell = random.choice(gameCells)
    wsValues[randomCell] = game.getNewValue()

    # Save and update values
    wb.save("values.xlsx")
    game.updateValues()


def move(direction):
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wb.active
    scoreAdded = 0

    # Game over text
    if wsValues['A2'].value != wsValues['B2'].value and wsValues['B2'].value != wsValues['C2'].value and \
            wsValues['C2'].value != wsValues['D2'].value and wsValues['A3'].value != wsValues['B3'].value and \
            wsValues['B3'].value != wsValues['C3'].value and wsValues['C3'].value != wsValues['D3'].value and \
            wsValues['A4'].value != wsValues['B4'].value and wsValues['B4'].value != wsValues['C4'].value and \
            wsValues['C4'].value != wsValues['D4'].value and wsValues['A2'].value != wsValues['A3'].value and \
            wsValues['A3'].value != wsValues['A4'].value and wsValues['A4'].value != wsValues['A5'].value and \
            wsValues['B2'].value != wsValues['B3'].value and wsValues['B3'].value != wsValues['B4'].value and \
            wsValues['B4'].value != wsValues['B5'].value and wsValues['C2'].value != wsValues['C3'].value and \
            wsValues['C3'].value != wsValues['C4'].value and wsValues['C4'].value != wsValues['C5'].value and \
            wsValues['D2'].value != wsValues['D3'].value and wsValues['D3'].value != wsValues['D4'].value and \
            wsValues['D4'].value != wsValues['D5'].value and wsValues['A2'] > 0 and wsValues['A3'] > 0 and \
            wsValues['A4'].value > 0 and wsValues['A5'].value > 0 and wsValues['B2'].value > 0 and \
            wsValues['B3'].value > 0 and wsValues['B4'].value > 0 and wsValues['B5'].value > 0 and \
            wsValues['C2'].value > 0 and wsValues['C3'].value > 0 and wsValues['C4'].value > 0 and \
            wsValues['C5'].value > 0 and wsValues['D2'].value > 0 and wsValues['D3'].value > 0 and \
            wsValues['D4'].value > 0 and wsValues['D5'].value > 0:
        wsValues['A6'] = "GAME OVER!"

    # Game won text
    if wsValues['A2'].value > 2000 or wsValues['A3'].value > 2000 or wsValues['A4'].value > 2000 or \
            wsValues['A5'].value > 2000 or wsValues['B2'].value > 2000 or wsValues['B3'].value > 2000 or \
            wsValues['B4'].value > 2000 or wsValues['B5'].value > 2000 or wsValues['C2'].value > 2000 or \
            wsValues['C3'].value > 2000 or wsValues['C4'].value > 2000 or wsValues['C5'].value > 2000 or \
            wsValues['D2'].value > 2000 or wsValues['D3'].value > 2000 or wsValues['D4'].value > 2000 or \
            wsValues['D5'].value > 2000:
        wsValues['A6'] = "YOU WON!"

    # Moves
    if direction == "up":
        if wsValues['A2'].value == 0 and wsValues['A3'].value == 0 and wsValues['A4'].value == 0:
            wsValues['A2'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A2'].value == 0 and wsValues['A3'].value == 0:
            wsValues['A2'] = wsValues['A4'].value
            wsValues['A3'] = wsValues['A5'].value
            wsValues['A4'] = 0
            wsValues['A5'] = 0
        if wsValues['A2'].value == 0:
            wsValues['A2'] = wsValues['A3'].value
            wsValues['A3'] = wsValues['A4'].value
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A3'].value == 0 and wsValues['A4'].value == 0:
            wsValues['A3'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A3'].value == 0:
            wsValues['A3'] = wsValues['A4'].value
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A4'].value == 0:
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['B2'].value == 0 and wsValues['B3'].value == 0 and wsValues['B4'].value == 0:
            wsValues['B2'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B2'].value == 0 and wsValues['B3'].value == 0:
            wsValues['B2'] = wsValues['B4'].value
            wsValues['B3'] = wsValues['B5'].value
            wsValues['B4'] = 0
            wsValues['B5'] = 0
        if wsValues['B2'].value == 0:
            wsValues['B2'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B3'].value == 0 and wsValues['B4'].value == 0:
            wsValues['B3'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B3'].value == 0:
            wsValues['B3'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B4'].value == 0:
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['C2'].value == 0 and wsValues['C3'].value == 0 and wsValues['C4'].value == 0:
            wsValues['C2'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C2'].value == 0 and wsValues['C3'].value == 0:
            wsValues['C2'] = wsValues['C4'].value
            wsValues['C3'] = wsValues['C5'].value
            wsValues['C4'] = 0
            wsValues['C5'] = 0
        if wsValues['C2'].value == 0:
            wsValues['C2'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C3'].value == 0 and wsValues['C4'].value == 0:
            wsValues['C3'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C3'].value == 0:
            wsValues['C3'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C4'].value == 0:
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['D2'].value == 0 and wsValues['D3'].value == 0 and wsValues['D4'].value == 0:
            wsValues['D2'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D2'].value == 0 and wsValues['D3'].value == 0:
            wsValues['D2'] = wsValues['D4'].value
            wsValues['D3'] = wsValues['D5'].value
            wsValues['D4'] = 0
            wsValues['D5'] = 0
        if wsValues['D2'].value == 0:
            wsValues['D2'] = wsValues['D3'].value
            wsValues['D3'] = wsValues['D4'].value
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D3'].value == 0 and wsValues['D4'].value == 0:
            wsValues['D3'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D3'].value == 0:
            wsValues['D3'] = wsValues['D4'].value
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D4'].value == 0:
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['A2'].value == wsValues['A3'].value:
            wsValues['A2'] = wsValues['A3'].value * 2
            scoreAdded += wsValues['A2'].value
            wsValues['A3'] = wsValues['A4'].value
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A3'].value == wsValues['A4'].value:
            wsValues['A3'] = wsValues['A4'].value * 2
            scoreAdded += wsValues['A3'].value
            wsValues['A4'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['A4'].value == wsValues['A5'].value:
            wsValues['A4'] = wsValues['A5'].value * 2
            scoreAdded += wsValues['A4'].value
            wsValues['A5'] = 0
        if wsValues['B2'].value == wsValues['B3'].value:
            wsValues['B2'] = wsValues['B3'].value * 2
            scoreAdded += wsValues['B2'].value
            wsValues['B3'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B3'].value == wsValues['B4'].value:
            wsValues['B3'] = wsValues['B4'].value * 2
            scoreAdded += wsValues['B3'].value
            wsValues['B4'] = wsValues['B5'].value
            wsValues['B5'] = 0
        if wsValues['B4'].value == wsValues['B5'].value:
            wsValues['B4'] = wsValues['B5'].value * 2
            scoreAdded += wsValues['B4'].value
            wsValues['B5'] = 0
        if wsValues['C2'].value == wsValues['C3'].value:
            wsValues['C2'] = wsValues['C3'].value * 2
            scoreAdded += wsValues['C2'].value
            wsValues['C3'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C3'].value == wsValues['C4'].value:
            wsValues['C3'] = wsValues['C4'].value * 2
            scoreAdded += wsValues['C3'].value
            wsValues['C4'] = wsValues['C5'].value
            wsValues['C5'] = 0
        if wsValues['C4'].value == wsValues['C5'].value:
            wsValues['C4'] = wsValues['C5'].value * 2
            scoreAdded += wsValues['C4'].value
            wsValues['C5'] = 0
        if wsValues['D2'].value == wsValues['D3'].value:
            wsValues['D2'] = wsValues['D3'].value * 2
            scoreAdded += wsValues['D2'].value
            wsValues['D3'] = wsValues['D4'].value
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D3'].value == wsValues['D4'].value:
            wsValues['D3'] = wsValues['D4'].value * 2
            scoreAdded += wsValues['D3'].value
            wsValues['D4'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['D4'].value == wsValues['D5'].value:
            wsValues['D4'] = wsValues['D5'].value * 2
            scoreAdded += wsValues['D4'].value
            wsValues['D5'] = 0

    elif direction == "down":
        if wsValues['A5'].value == 0 and wsValues['A4'].value == 0 and wsValues['A3'].value == 0:
            wsValues['A5'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A5'].value == 0 and wsValues['A4'].value == 0:
            wsValues['A5'] = wsValues['A3'].value
            wsValues['A4'] = wsValues['A2'].value
            wsValues['A3'] = 0
            wsValues['A2'] = 0
        if wsValues['A5'].value == 0:
            wsValues['A5'] = wsValues['A4'].value
            wsValues['A4'] = wsValues['A3'].value
            wsValues['A3'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A4'].value == 0 and wsValues['A3'].value == 0:
            wsValues['A4'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A4'].value == 0:
            wsValues['A4'] = wsValues['A3'].value
            wsValues['A3'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A3'].value == 0:
            wsValues['A3'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['B5'].value == 0 and wsValues['B4'].value == 0 and wsValues['B3'].value == 0:
            wsValues['B5'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B5'].value == 0 and wsValues['B4'].value == 0:
            wsValues['B5'] = wsValues['B3'].value
            wsValues['B4'] = wsValues['B2'].value
            wsValues['B3'] = 0
            wsValues['B2'] = 0
        if wsValues['B5'].value == 0:
            wsValues['B5'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B4'].value == 0 and wsValues['B3'].value == 0:
            wsValues['B4'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B4'].value == 0:
            wsValues['B4'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B3'].value == 0:
            wsValues['B3'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['C5'].value == 0 and wsValues['C4'].value == 0 and wsValues['C3'].value == 0:
            wsValues['C5'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C5'].value == 0 and wsValues['C4'].value == 0:
            wsValues['C5'] = wsValues['C3'].value
            wsValues['C4'] = wsValues['C2'].value
            wsValues['C3'] = 0
            wsValues['C2'] = 0
        if wsValues['C5'].value == 0:
            wsValues['C5'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C4'].value == 0 and wsValues['C3'].value == 0:
            wsValues['C4'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C4'].value == 0:
            wsValues['C4'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C3'].value == 0:
            wsValues['C3'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['D5'].value == 0 and wsValues['D4'].value == 0 and wsValues['D3'].value == 0:
            wsValues['D5'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D5'].value == 0 and wsValues['D4'].value == 0:
            wsValues['D5'] = wsValues['D3'].value
            wsValues['D4'] = wsValues['D2'].value
            wsValues['D3'] = 0
            wsValues['D2'] = 0
        if wsValues['D5'].value == 0:
            wsValues['D5'] = wsValues['D4'].value
            wsValues['D4'] = wsValues['D3'].value
            wsValues['D3'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D4'].value == 0 and wsValues['D3'].value == 0:
            wsValues['D4'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D4'].value == 0:
            wsValues['D4'] = wsValues['D3'].value
            wsValues['D3'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D3'].value == 0:
            wsValues['D3'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['A5'].value == wsValues['A4'].value:
            wsValues['A5'] = wsValues['A4'].value * 2
            scoreAdded += wsValues['A5'].value
            wsValues['A4'] = wsValues['A3'].value
            wsValues['A3'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A4'].value == wsValues['A3'].value:
            wsValues['A4'] = wsValues['A3'].value * 2
            scoreAdded += wsValues['A4'].value
            wsValues['A3'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['A3'].value == wsValues['A2'].value:
            wsValues['A3'] = wsValues['A2'].value * 2
            scoreAdded += wsValues['A3'].value
            wsValues['A2'] = 0
        if wsValues['B5'].value == wsValues['B4'].value:
            wsValues['B5'] = wsValues['B4'].value * 2
            scoreAdded += wsValues['B5'].value
            wsValues['B4'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B4'].value == wsValues['B3'].value:
            wsValues['B4'] = wsValues['B3'].value * 2
            scoreAdded += wsValues['B4'].value
            wsValues['B3'] = wsValues['B2'].value
            wsValues['B2'] = 0
        if wsValues['B3'].value == wsValues['B2'].value:
            wsValues['B3'] = wsValues['B2'].value * 2
            scoreAdded += wsValues['B3'].value
            wsValues['B2'] = 0
        if wsValues['C5'].value == wsValues['C4'].value:
            wsValues['C5'] = wsValues['C4'].value * 2
            scoreAdded += wsValues['C5'].value
            wsValues['C4'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C4'].value == wsValues['C3'].value:
            wsValues['C4'] = wsValues['C3'].value * 2
            scoreAdded += wsValues['C4'].value
            wsValues['C3'] = wsValues['C2'].value
            wsValues['C2'] = 0
        if wsValues['C3'].value == wsValues['C2'].value:
            wsValues['C3'] = wsValues['C2'].value * 2
            scoreAdded += wsValues['C3'].value
            wsValues['C2'] = 0
        if wsValues['D5'].value == wsValues['D4'].value:
            wsValues['D5'] = wsValues['D4'].value * 2
            scoreAdded += wsValues['D5'].value
            wsValues['D4'] = wsValues['D3'].value
            wsValues['D3'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D4'].value == wsValues['D3'].value:
            wsValues['D4'] = wsValues['D3'].value * 2
            scoreAdded += wsValues['D4'].value
            wsValues['D3'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['D3'].value == wsValues['D2'].value:
            wsValues['D3'] = wsValues['D2'].value * 2
            scoreAdded += wsValues['D3'].value
            wsValues['D2'] = 0

    elif direction == "left":
        if wsValues['A2'].value == 0 and wsValues['B2'].value == 0 and wsValues['C2'].value == 0:
            wsValues['A2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['A2'].value == 0 and wsValues['B2'].value == 0:
            wsValues['A2'] = wsValues['C2'].value
            wsValues['B2'] = wsValues['D2'].value
            wsValues['C2'] = 0
            wsValues['D2'] = 0
        if wsValues['A2'].value == 0:
            wsValues['A2'] = wsValues['B2'].value
            wsValues['B2'] = wsValues['C2'].value
            wsValues['C2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['B2'].value == 0 and wsValues['C2'].value == 0:
            wsValues['B2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['B2'].value == 0:
            wsValues['B2'] = wsValues['C2'].value
            wsValues['C2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['C2'].value == 0:
            wsValues['C2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['A3'].value == 0 and wsValues['B3'].value == 0 and wsValues['C3'].value == 0:
            wsValues['A3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['A3'].value == 0 and wsValues['B3'].value == 0:
            wsValues['A3'] = wsValues['C3'].value
            wsValues['B3'] = wsValues['D3'].value
            wsValues['C3'] = 0
            wsValues['D3'] = 0
        if wsValues['A3'].value == 0:
            wsValues['A3'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['B3'].value == 0 and wsValues['C3'].value == 0:
            wsValues['B3'] = wsValues['D2'].value
            wsValues['D3'] = 0
        if wsValues['B3'].value == 0:
            wsValues['B3'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['C3'].value == 0:
            wsValues['C3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['A4'].value == 0 and wsValues['B4'].value == 0 and wsValues['C4'].value == 0:
            wsValues['A4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['A4'].value == 0 and wsValues['B4'].value == 0:
            wsValues['A4'] = wsValues['C4'].value
            wsValues['B4'] = wsValues['D4'].value
            wsValues['C4'] = 0
            wsValues['D4'] = 0
        if wsValues['A4'].value == 0:
            wsValues['A4'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['B4'].value == 0 and wsValues['C4'].value == 0:
            wsValues['B4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['B4'].value == 0:
            wsValues['B4'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['C4'].value == 0:
            wsValues['C4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['A5'].value == 0 and wsValues['B5'].value == 0 and wsValues['C5'].value == 0:
            wsValues['A5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['A5'].value == 0 and wsValues['B5'].value == 0:
            wsValues['A5'] = wsValues['C5'].value
            wsValues['B5'] = wsValues['D5'].value
            wsValues['C5'] = 0
            wsValues['D5'] = 0
        if wsValues['A5'].value == 0:
            wsValues['A5'] = wsValues['B5'].value
            wsValues['B5'] = wsValues['C5'].value
            wsValues['C5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['B5'].value == 0 and wsValues['C5'].value == 0:
            wsValues['B5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['B5'].value == 0:
            wsValues['B5'] = wsValues['C5'].value
            wsValues['C5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['C5'].value == 0:
            wsValues['C5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['A2'].value == wsValues['B2'].value:
            wsValues['A2'] = wsValues['B2'].value * 2
            scoreAdded += wsValues['A2'].value
            wsValues['B2'] = wsValues['C2'].value
            wsValues['C2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['B2'].value == wsValues['C2'].value:
            wsValues['B2'] = wsValues['C2'].value * 2
            scoreAdded += wsValues['B2'].value
            wsValues['C2'] = wsValues['D2'].value
            wsValues['D2'] = 0
        if wsValues['C2'].value == wsValues['D2'].value:
            wsValues['C2'] = wsValues['D2'].value * 2
            scoreAdded += wsValues['C2'].value
            wsValues['D2'] = 0
        if wsValues['A3'].value == wsValues['B3'].value:
            wsValues['A3'] = wsValues['B3'].value * 2
            scoreAdded += wsValues['A3'].value
            wsValues['B3'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['B3'].value == wsValues['C3'].value:
            wsValues['B3'] = wsValues['C3'].value * 2
            scoreAdded += wsValues['B3'].value
            wsValues['C3'] = wsValues['D3'].value
            wsValues['D3'] = 0
        if wsValues['C3'].value == wsValues['D3'].value:
            wsValues['C3'] = wsValues['D3'].value * 2
            scoreAdded += wsValues['C3'].value
            wsValues['D3'] = 0
        if wsValues['A4'].value == wsValues['B4'].value:
            wsValues['A4'] = wsValues['B4'].value * 2
            scoreAdded += wsValues['A4'].value
            wsValues['B4'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['B4'].value == wsValues['C4'].value:
            wsValues['B4'] = wsValues['C4'].value * 2
            scoreAdded += wsValues['B4'].value
            wsValues['C4'] = wsValues['D4'].value
            wsValues['D4'] = 0
        if wsValues['C4'].value == wsValues['D4'].value:
            wsValues['C4'] = wsValues['D4'].value * 2
            scoreAdded += wsValues['C4'].value
            wsValues['D4'] = 0
        if wsValues['A5'].value == wsValues['B5'].value:
            wsValues['A5'] = wsValues['B5'].value * 2
            scoreAdded += wsValues['A5'].value
            wsValues['B5'] = wsValues['C5'].value
            wsValues['C5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['B5'].value == wsValues['C5'].value:
            wsValues['B5'] = wsValues['C5'].value * 2
            scoreAdded += wsValues['B5'].value
            wsValues['C5'] = wsValues['D5'].value
            wsValues['D5'] = 0
        if wsValues['C5'].value == wsValues['D5'].value:
            wsValues['C5'] = wsValues['D5'].value * 2
            scoreAdded += wsValues['C5'].value
            wsValues['D5'] = 0

    elif direction == "right":
        if wsValues['D2'].value == 0 and wsValues['C2'].value == 0 and wsValues['B2'].value == 0:
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['D2'].value == 0 and wsValues['C2'].value == 0:
            wsValues['D2'] = wsValues['B2'].value
            wsValues['C2'] = wsValues['A2'].value
            wsValues['B2'] = 0
            wsValues['A2'] = 0
        if wsValues['D2'].value == 0:
            wsValues['D2'] = wsValues['C2'].value
            wsValues['C2'] = wsValues['B2'].value
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['C2'].value == 0 and wsValues['B2'].value == 0:
            wsValues['C2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['C2'].value == 0:
            wsValues['C2'] = wsValues['B2'].value
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['B2'].value == 0:
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['D3'].value == 0 and wsValues['C3'].value == 0 and wsValues['B3'].value == 0:
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['D3'].value == 0 and wsValues['C3'].value == 0:
            wsValues['D3'] = wsValues['B3'].value
            wsValues['C3'] = wsValues['A3'].value
            wsValues['B3'] = 0
            wsValues['A3'] = 0
        if wsValues['D3'].value == 0:
            wsValues['D3'] = wsValues['C3'].value
            wsValues['C3'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['C3'].value == 0 and wsValues['B3'].value == 0:
            wsValues['C3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['C3'].value == 0:
            wsValues['C3'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['B3'].value == 0:
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['D4'].value == 0 and wsValues['C4'].value == 0 and wsValues['B4'].value == 0:
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['D4'].value == 0 and wsValues['C4'].value == 0:
            wsValues['D4'] = wsValues['B4'].value
            wsValues['C4'] = wsValues['A4'].value
            wsValues['B4'] = 0
            wsValues['A4'] = 0
        if wsValues['D4'].value == 0:
            wsValues['D4'] = wsValues['C4'].value
            wsValues['C4'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['C4'].value == 0 and wsValues['B4'].value == 0:
            wsValues['C4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['C4'].value == 0:
            wsValues['C4'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['B4'].value == 0:
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['D5'].value == 0 and wsValues['C5'].value == 0 and wsValues['B5'].value == 0:
            wsValues['B5'] = wsValues['A2'].value
            wsValues['A5'] = 0
        if wsValues['D5'].value == 0 and wsValues['C5'].value == 0:
            wsValues['D5'] = wsValues['B5'].value
            wsValues['C5'] = wsValues['A5'].value
            wsValues['B5'] = 0
            wsValues['A5'] = 0
        if wsValues['D5'].value == 0:
            wsValues['D5'] = wsValues['C5'].value
            wsValues['C5'] = wsValues['B5'].value
            wsValues['B5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['C5'].value == 0 and wsValues['B5'].value == 0:
            wsValues['C5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['C5'].value == 0:
            wsValues['C5'] = wsValues['B5'].value
            wsValues['B5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['B5'].value == 0:
            wsValues['B5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['D2'].value == wsValues['C2'].value:
            wsValues['D2'] = wsValues['C2'].value * 2
            scoreAdded += wsValues['D2'].value
            wsValues['C2'] = wsValues['B2'].value
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['C2'].value == wsValues['B2'].value:
            wsValues['C2'] = wsValues['B2'].value * 2
            scoreAdded += wsValues['C2'].value
            wsValues['B2'] = wsValues['A2'].value
            wsValues['A2'] = 0
        if wsValues['B2'].value == wsValues['A2'].value:
            wsValues['B2'] = wsValues['A2'].value * 2
            scoreAdded += wsValues['B2'].value
            wsValues['A2'] = 0
        if wsValues['D3'].value == wsValues['C3'].value:
            wsValues['D3'] = wsValues['C3'].value * 2
            scoreAdded += wsValues['D3'].value
            wsValues['C3'] = wsValues['B3'].value
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['C3'].value == wsValues['B3'].value:
            wsValues['C3'] = wsValues['B3'].value * 2
            scoreAdded += wsValues['C3'].value
            wsValues['B3'] = wsValues['A3'].value
            wsValues['A3'] = 0
        if wsValues['B3'].value == wsValues['A3'].value:
            wsValues['B3'] = wsValues['A3'].value * 2
            scoreAdded += wsValues['B3'].value
            wsValues['A3'] = 0
        if wsValues['D4'].value == wsValues['C4'].value:
            wsValues['D4'] = wsValues['C4'].value * 2
            scoreAdded += wsValues['D4'].value
            wsValues['C4'] = wsValues['B4'].value
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['C4'].value == wsValues['B4'].value:
            wsValues['C4'] = wsValues['B4'].value * 2
            scoreAdded += wsValues['C4'].value
            wsValues['B4'] = wsValues['A4'].value
            wsValues['A4'] = 0
        if wsValues['B4'].value == wsValues['A4'].value:
            wsValues['B4'] = wsValues['A4'].value * 2
            scoreAdded += wsValues['B4'].value
            wsValues['A4'] = 0
        if wsValues['D5'].value == wsValues['C5'].value:
            wsValues['D5'] = wsValues['C5'].value * 2
            scoreAdded += wsValues['D5'].value
            wsValues['C5'] = wsValues['B5'].value
            wsValues['B5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['C5'].value == wsValues['B5'].value:
            wsValues['C5'] = wsValues['B5'].value * 2
            scoreAdded += wsValues['C5'].value
            wsValues['B5'] = wsValues['A5'].value
            wsValues['A5'] = 0
        if wsValues['B5'].value == wsValues['A5'].value:
            wsValues['B5'] = wsValues['A5'].value * 2
            scoreAdded += wsValues['B5'].value
            wsValues['A5'] = 0

    # Generate new block in an empty cell
    emptyCells = ['A2', 'A3', 'A4', 'A5', 'B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5']
    for i in emptyCells:
        if wsValues[i].value > 0:
            emptyCells.remove(i)
    newBlock = random.choice(emptyCells)
    wsValues[newBlock] = game.getNewValue()

    # Save, update score and values
    wb.save("values.xlsx")
    game.updateValues()
    game.updateScore(scoreAdded)
