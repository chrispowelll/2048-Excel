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

    # Rest game status
    wsValues['A6'] = ""

    # Reset moves made
    wsValues['D1'] = 0

    # Create random starting block
    randomCell = random.choice(gameCells)
    wsValues[randomCell] = game.getNewValue()

    # Save and update values
    wb.save("values.xlsx")


def move(direction):
    wb = openpyxl.load_workbook("values.xlsx", read_only=False)
    wsValues = wb.active

    # Get current values to check if move is made later
    originalValues = []
    for i in ['A', 'B', 'C', 'D']:
        for j in ['2', '3', '4', '5']:
            cell = wsValues[str(i) + str(j)].value
            originalValues.append(cell)

    # Moves
    for c in ['A', 'B', 'C', 'D']:
        if direction == "up":
            # Fill empty cells above
            if wsValues[str(c) + str(4)].value == 0:
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value
                wsValues[str(c) + str(5)] = 0
            if wsValues[str(c) + str(3)].value == 0:
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(4)].value
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value
                wsValues[str(c) + str(5)] = 0
            if wsValues[str(c) + str(2)].value == 0:
                wsValues[str(c) + str(2)] = wsValues[str(c) + str(3)].value
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(4)].value
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value
                wsValues[str(c) + str(5)] = 0

            # Complete matches and add scores
            if wsValues[str(c) + str(2)].value == wsValues[str(c) + str(3)].value:
                wsValues[str(c) + str(2)] = wsValues[str(c) + str(3)].value * 2
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(4)].value
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value
                wsValues[str(c) + str(5)] = 0
            if wsValues[str(c) + str(3)].value == wsValues[str(c) + str(4)].value:
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(4)].value * 2
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value
                wsValues[str(c) + str(5)] = 0
            if wsValues[str(c) + str(4)].value == wsValues[str(c) + str(5)].value:
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(5)].value * 2
                wsValues[str(c) + str(5)] = 0

        elif direction == "down":
            # Fill empty cells below
            if wsValues[str(c) + str(3)].value == 0:
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value
                wsValues[str(c) + str(2)] = 0
            if wsValues[str(c) + str(4)].value == 0:
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(3)].value
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value
                wsValues[str(c) + str(2)] = 0
            if wsValues[str(c) + str(5)].value == 0:
                wsValues[str(c) + str(5)] = wsValues[str(c) + str(4)].value
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(3)].value
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value
                wsValues[str(c) + str(2)] = 0

            # Complete matches and add scores
            if wsValues[str(c) + str(5)].value == wsValues[str(c) + str(4)].value:
                wsValues[str(c) + str(5)] = wsValues[str(c) + str(4)].value * 2
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(3)].value
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value
                wsValues[str(c) + str(2)] = 0
            if wsValues[str(c) + str(4)].value == wsValues[str(c) + str(3)].value:
                wsValues[str(c) + str(4)] = wsValues[str(c) + str(3)].value * 2
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value
                wsValues[str(c) + str(2)] = 0
            if wsValues[str(c) + str(3)].value == wsValues[str(c) + str(2)].value:
                wsValues[str(c) + str(3)] = wsValues[str(c) + str(2)].value * 2
                wsValues[str(c) + str(2)] = 0

    for r in ['2', '3', '4', '5']:
        if direction == "left":
            # Fill empty cells to the left
            if wsValues["c" + str(r)].value == 0:
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value
                wsValues["d" + str(r)] = 0
            if wsValues["b" + str(r)].value == 0:
                wsValues["b" + str(r)] = wsValues["c" + str(r)].value
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value
                wsValues["d" + str(r)] = 0
            if wsValues["a" + str(r)].value == 0:
                wsValues["a" + str(r)] = wsValues["b" + str(r)].value
                wsValues["b" + str(r)] = wsValues["c" + str(r)].value
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value
                wsValues["d" + str(r)] = 0

            # Complete matches and add scores
            if wsValues["a" + str(r)].value == wsValues["b" + str(r)].value:
                wsValues["a" + str(r)] = wsValues["b" + str(r)].value * 2
                wsValues["b" + str(r)] = wsValues["c" + str(r)].value
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value
                wsValues["d" + str(r)] = 0
            if wsValues["b" + str(r)].value == wsValues["c" + str(r)].value:
                wsValues["b" + str(r)] = wsValues["c" + str(r)].value * 2
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value
                wsValues["d" + str(r)] = 0
            if wsValues["c" + str(r)].value == wsValues["d" + str(r)].value:
                wsValues["c" + str(r)] = wsValues["d" + str(r)].value * 2
                wsValues["d" + str(r)] = 0

        elif direction == "right":
            # Fill empty cells to the right
            if wsValues["b" + str(r)].value == 0:
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value
                wsValues["a" + str(r)] = 0
            if wsValues["c" + str(r)].value == 0:
                wsValues["c" + str(r)] = wsValues["b" + str(r)].value
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value
                wsValues["a" + str(r)] = 0
            if wsValues["d" + str(r)].value == 0:
                wsValues["d" + str(r)] = wsValues["c" + str(r)].value
                wsValues["c" + str(r)] = wsValues["b" + str(r)].value
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value
                wsValues["a" + str(r)] = 0

            # Complete matches and add scores
            if wsValues["d" + str(r)].value == wsValues["c" + str(r)].value:
                wsValues["d" + str(r)] = wsValues["c" + str(r)].value * 2
                wsValues["c" + str(r)] = wsValues["b" + str(r)].value
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value
                wsValues["a" + str(r)] = 0
            if wsValues["c" + str(r)].value == wsValues["b" + str(r)].value:
                wsValues["c" + str(r)] = wsValues["b" + str(r)].value * 2
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value
                wsValues["a" + str(r)] = 0
            if wsValues["b" + str(r)].value == wsValues["a" + str(r)].value:
                wsValues["b" + str(r)] = wsValues["a" + str(r)].value * 2
                wsValues["a" + str(r)] = 0

    # Get current values to check if move was made
    updatedValues = []
    for i in ['A', 'B', 'C', 'D']:
        for j in ['2', '3', '4', '5']:
            cell = wsValues[str(i) + str(j)].value
            updatedValues.append(cell)

    # Check if move was made
    if updatedValues != originalValues:
        # Update moves made
        wsValues['D1'].value += 1

        # Generate list of empty cells
        emptyCells = []
        for i in ['A', 'B', 'C', 'D']:
            for j in ['2', '3', '4', '5']:
                cell = str(i) + str(j)
                if wsValues[cell].value == 0:
                    emptyCells.append(cell)

        # Create a new number of there is an empty cell
        if len(emptyCells) > 0:
            newBlock = random.choice(emptyCells)
            wsValues[newBlock] = game.getNewValue()
            emptyCells.remove(newBlock)

        # Detect if game is over
        if len(emptyCells) == 0 and wsValues['A2'].value != wsValues['A3'].value and wsValues['A2'].value != \
                wsValues['B2'].value and wsValues['B2'].value != wsValues['C2'].value and wsValues['B2'].value != \
                wsValues['B3'] and wsValues['C2'].value != wsValues['C3'].value and wsValues['C2'].value != \
                wsValues['D2'].value and wsValues['D2'].value != wsValues['D3'].value and wsValues['A3'].value != \
                wsValues['A4'].value and wsValues['A3'].value != wsValues['B3'].value and wsValues['B3'].value != \
                wsValues['B4'].value and wsValues['B3'].value != wsValues['C3'].value and wsValues['C3'].value != \
                wsValues['C4'].value and wsValues['C3'].value != wsValues['D3'].value and wsValues['D3'].value != \
                wsValues['D4'].value and wsValues['A4'].value != wsValues['A5'].value and wsValues['A4'].value != \
                wsValues['B4'].value and wsValues['B4'].value != wsValues['B5'].value and wsValues['B4'].value != \
                wsValues['C4'].value and wsValues['C4'].value != wsValues['C5'].value and wsValues['C4'].value != \
                wsValues['D4'].value and wsValues['D5'].value != wsValues['C5'].value and wsValues['C5'].value != \
                wsValues['B5'].value and wsValues['B5'].value != wsValues['A5'].value:
            # End game
            wsValues['A6'] = "GAME OVER!"

        # Detect if game won
        for i in ['A', 'B', 'C', 'D']:
            for j in ['2', '3', '4', '5']:
                if wsValues[str(i) + str(j)].value == 2048:
                    wsValues['A6'] = "GAME WON!"

    # Save, update score and values
    wb.save("values.xlsx")
    game.updateValues()
