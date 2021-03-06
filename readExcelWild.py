# Imports:
import random
import openpyxl
from openpyxl import DEFUSEDXML

# Declarations of global variables:
workbook = openpyxl.load_workbook(filename="customDiceOutcomes.xlsx")
diceSheet = workbook.active
settingSheet = workbook["settings"]

# Setup:
if DEFUSEDXML is False:
    print('You are at risk: Please install and setup DefusedXML.')

# Function definitions:
# A function which does exactly as the name implies.
def str_to_int_or_float(value):
    if isinstance(value, bool):
        return value
    try:
        return int(value)
    except ValueError and TypeError:
        try:
            return float(value)
        except ValueError and TypeError:
            return value

# Class definitions:
# Defines the table, with outputs and table-wide min/max-values.
class Table:
    actualTableNumber = 0
    startRow = 1

    minMaxRow = startRow + 1
    min = 1
    max = 100

    diceRoll = 1

    outcomeAmount = 0
    cellNumber = 0
    i = 0

    gotoTable = 0

    empty = 0
    def __init__(self, tableNumber):
        self.tableNumber = tableNumber

        # Validates table number: Should not be smaller than 1.
        if 0 < str_to_int_or_float(self.tableNumber):
            # The following loop makes the startRow value fit with the desired table.
            while self.actualTableNumber < self.tableNumber-1:
                while self.empty < 2:
                    if not diceSheet["B" + str(self.startRow)].value:
                        self.empty += 1
                        # print("Empty = " + str(self.empty))
                    else:
                        self.empty = 0
                        # print("Empty = " + str(self.empty))
                    self.startRow += 1
                self.actualTableNumber += 1
                self.empty = 0

            # Readying variables for upcoming loop
            self.minMaxRow = self.startRow + 1
            self.min = str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value)
            self.max = str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value)
            # This loop establishes the min and max values for the table.
            while self.empty < 2:
                if not diceSheet["B" + str(self.minMaxRow)].value:
                    self.empty += 1
                elif str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value) < self.min:
                    self.empty = 0
                    self.min = str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value)
                elif str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value) > self.max:
                    self.empty = 0
                    self.max = str_to_int_or_float(diceSheet["B" + str(self.minMaxRow)].value)
                else:  # Because str_to_int_or_float does not give errors when its input isn't
                    # compatible (letters, for instance).
                    self.empty = 0
                self.minMaxRow += 1

            # Array to store possible outcomes in
            self.OutcomeArray = []

            # Create list of possible outcomes:
            self.outcomeAmount = int(((self.minMaxRow - 4) - (self.startRow + 1)) / 3 + 1)  # It makes sense.
            # print(str(self.outcomeAmount))
            self.cellNumber = self.startRow + 1
            while self.i < self.outcomeAmount:
                self.outcome = TableContent(self.cellNumber)
                self.OutcomeArray.append(self.outcome)

                self.i += 1
                self.cellNumber += 3
        else:
            print("Minor error: Invalid table number. Please use a valid table number in accordance with the README.txt")

        self.exceptMin = 0
        self.exceptMax = 1

    def getTableInfo(self):
        print("Table no. " + str(self.tableNumber) + " begins at row " + str(self.startRow) + ".")
        print("It has " + str(self.min) + " as min, and " + str(self.max) + " as max.")
        print("It has " + str(self.outcomeAmount) + " outcomes:")
        for n in range(0, self.outcomeAmount):
            self.OutcomeArray[n].getContentInfo()
        print("The first possible outcome has " + str(self.OutcomeArray[0].min) + " as min, and " + str(self.OutcomeArray[0].max) + " as max.")

    def roll(self):
        # Utilizes the Random module, to "throw the dice" between Min and Max.
        self.diceRoll = random.randint(self.min, self.max)
        print("You rolled " + str(self.diceRoll))
        print('')

        # Finds the first outcome which the rolled value is within.
        for n in range(0, self.outcomeAmount):
            if self.OutcomeArray[n].min <= self.diceRoll <= self.OutcomeArray[n].max:
                self.gotoTable, self.exceptMin, self.exceptMax = self.OutcomeArray[n].getContentInfo()

        return self.gotoTable, self.exceptMin, self.exceptMax

    def gotoRoll(self, exceptMin = 0, exceptMax = 0):
        self.exceptMin = exceptMin
        self.exceptMax = exceptMax

        # Utilizes the Random module to "throw the dice" between Min and Max, discards results within the except limits.
        self.diceRoll = random.randint(self.min, self.max)
        if self.exceptMin <= self.diceRoll <= self.exceptMax:
            while self.exceptMin <= self.diceRoll <= self.exceptMax:
                self.diceRoll = random.randint(self.min, self.max)
        print("You rolled " + str(self.diceRoll))
        print('')

        # Finds the first outcome which the rolled value is within.
        for n in range(0, self.outcomeAmount):
            if self.OutcomeArray[n].min <= self.diceRoll <= self.OutcomeArray[n].max:
                self.gotoTable, self.exceptMin, self.exceptMax = self.OutcomeArray[n].getContentInfo()

        return self.gotoTable, self.exceptMin, self.exceptMax


# Defines the possible outcomes, as well as their min/max-values.
class TableContent:
    min = 0
    max = 100
    output = "Empty"
    gotoTable = 0
    exceptMin = 0
    exceptMax = 0

    def __init__(self, row):
        self.row = row
        # Checks if the user has typed the min and max values into the correct cells. Otherwise it switches them.
        if str_to_int_or_float(diceSheet["B" + str(self.row)].value) < str_to_int_or_float(diceSheet["B" + str(self.row + 1)].value):
            self.min = str_to_int_or_float(diceSheet["B" + str(self.row)].value)
            self.max = str_to_int_or_float(diceSheet["B" + str(self.row + 1)].value)
        else:
            self.min = str_to_int_or_float(diceSheet["B" + str(self.row + 1)].value)
            self.max = str_to_int_or_float(diceSheet["B" + str(self.row)].value)

        # Reads the output
        self.output = diceSheet["C" + str(self.row)].value

        # Checks for possible values in the goto-column.
        if diceSheet["D" + str(self.row)].value:
            self.gotoTable = str_to_int_or_float(diceSheet["D" + str(self.row)].value)
            # Checks for possible values in the except-column
            if diceSheet["E" + str(self.row)].value and diceSheet["E" + str(self.row + 1)].value:
                # Checks if the user has typed the min/max values into the correct cells. Otherwise it switches them.
                # Prepare vars
                if str_to_int_or_float(diceSheet["E" + str(self.row)].value) < str_to_int_or_float(
                        diceSheet["E" + str(self.row + 1)].value):
                    self.exceptMin = str_to_int_or_float(diceSheet["E" + str(self.row)].value)
                    self.exceptMax = str_to_int_or_float(diceSheet["E" + str(self.row + 1)].value)
                else:
                    self.exceptMin = str_to_int_or_float(diceSheet["E" + str(self.row + 1)].value)
                    self.exceptMax = str_to_int_or_float(diceSheet["E" + str(self.row)].value)



    def getContentInfo(self):
        print(str(self.output))
        print('')
        return self.gotoTable, self.exceptMin, self.exceptMax


# Running code:
# Redying variables for the upcoming loop:
empty = 0
startRow = 1
tableAmount = 0
# The following loop makes the startRow value fit with the desired table.
while True:
    while empty < 2:
        if not diceSheet["B" + str(startRow)].value:
            empty += 1
            # print("Empty = " + str(self.empty))
        else:
            empty = 0
            # print("Empty = " + str(self.empty))
        startRow += 1
    tableAmount += 1
    if not diceSheet["B" + str(startRow)].value:
        break
    empty = 0

# Array to store tables in
TableArray = []

# Redying variables:
i = 0
# Storing tables in the appropriate array.
while i < tableAmount:
    ListedTable = Table(i + 1)
    TableArray.append(ListedTable)
    i += 1

# Following loop will accept input, and exit if the input is either not a valid type, or smaller than/equal to 0.
# Redying variables:
gotoTable = 0
while True:
    userInput = input('What table do you want to roll from? (integer): ')

    # Make sure that the input is of the correct type:
    while True:
        try:
            userInput = int(userInput)
            break
        except ValueError:
            while True:
                if userInput.isdigit():
                    break
                else:
                    try:
                        userInput = float(userInput)
                        if float(userInput) < float((int(userInput) + 0.5)):
                            userInput = int(float(userInput))
                        else:
                            userInput = int(float(userInput)) + 1
                        print()
                        print('Lets just call that ' + str(userInput) + '.')
                        print()
                        break
                    except ValueError:
                        print()
                        userInput = input('...in integers, please: ')

    # If input is greater than 0, the appropriate table will be found.
    if 0 < userInput <= tableAmount:
        print('Rolling on table ' + str(userInput) + '...')
        gotoTable, exceptMin, exceptMax = TableArray[userInput - 1].roll()
        recursiveProtectionVariable = 0

    elif len(TableArray) < userInput:
        print('There aren\'t that many tables to choose from.')
        print('Please write a number smaller than, or equal to ' + str(tableAmount) + '.')
        print('Or write a number smaller than 0, if you want to exit the program.')

    elif userInput <= 0:
        break

    # If a table number is stored in the 'GoTo'-cell, roll in that table as well.
    while gotoTable:
        print('Rolling on table ' + str(gotoTable) + '...')
        gotoTable, exceptMin, exceptMax = TableArray[gotoTable - 1].gotoRoll(exceptMin, exceptMax)
        # If the program has executed more than x tabels since user-input,
        # it will shut down, as to not run in an infinite loop.
        recursiveProtectionVariable += 1
        if str_to_int_or_float(settingSheet["B2"].value) < recursiveProtectionVariable:
            gotoTable = 0
            print('Execution stopped, due to tabels being executed more than 20 times without user-input.')


# Pretty useless piece of code, only here so the program won't run, and then quit on itself before the user gets a-
# chance to read the output.
print('')
exitCode = (input('Press Enter to exit.'))
