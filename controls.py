import openpyxl
import game
import random


def moveUp():
    print("move up")


def moveDown():
    print("move down")


def moveLeft():
    print("move left")


def moveRight():
    print("move right")


def startGame():
    wb = openpyxl.load_workbook("2048.xlsx")
    ws2048 = wb.active
    gameCells = ['A2', 'A3', 'A4', 'A5', 'B2', 'B3', 'B4', 'B5', 'C2', 'C3', 'C4', 'C5', 'D2', 'D3', 'D4', 'D5']

    # Fill board with 0s
    for i in gameCells:  # game
        ws2048[i] = 0

    # Create starting number and starting score
    ws2048['D1'] = "0"
    randomCell = random.choice(gameCells)
    ws2048[randomCell] = str(game.getNewValue())

    # Save
    wb.save("2048.xlsx")
