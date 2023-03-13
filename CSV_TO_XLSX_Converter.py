# This program converts csv to xlsx.

#!/usr/local/bin/python3.10
import pandas as pd # import pandas module to convert csv to xlsx
from openpyxl import load_workbook # import openpyxl to parse and modify csv


# function below removes excel column that has "Currency code"
def removeCurrencyColumn(excelFile):
    # print("path: " + excelFile)

    book = load_workbook(excelFile)
    sheet = book.active  # iterable

    # check where "Currency code" is & the column
    for row in sheet:
        for cell in row:
            try:
                if 'Currency code' in cell.value:
                    # print("Currency code is at: ", cell.coordinate, cell.column)
                    sheet.delete_cols(cell.column, 1) # delete id column

            except TypeError:
                continue

    book.save(excelFile) # saves  file

# function below changes csv to excel
def convertCSV2Excel(docName, excelFile):
    
    try:
        read_file = pd.read_csv (docName) # can handle filepath

        # convert to excel if no errors reading file
        try:
            read_file.to_excel (excelFile, index = None, header=False) # can handle filepath
            return True
        except UnboundLocalError:
            print("Error 2: Writing to file. Please make sure you have permission to write.")

    except pd.errors.ParserError:
        print("Error 1: reading file.")


# converter() function converts csv to xlsx then calls removeCurrencyColumn() to remove appropriate column
def converter(numFiles, docNames):
    num = 0
    while num < numFiles:
        docName = docNames[num]

        docNameLastOnly = docName.split("/")

        if "csv" in docName: # checks if provided file is csv
            excelFile = docName.replace("csv", "xlsx")
            # if convertCSV2Excel(docName, excelFile): # converts to xlsx & returns true
            #     removeCurrencyColumn(excelFile) # removes currency column
            # else:
            #     print("There was an issue with the" + docNameLastOnly[len(docNameLastOnly)-1] + ", please try a different one.")
            convertCSV2Excel(docName, excelFile)
        else:
            # removeCurrencyColumn(excelFile) # if not csv, removes column
            print("please enter a csv file")

        num += 1

    print("done")

# function below receives file names
def bulkUpdater():

    print("Welcome! To enter multiple documents, paste in TextEdit, then add comma at end")
    print("_" * 30)
    docName = input("Enter file names: ")
    docName = docName.replace(".csv /", ".csv,/")
    print(docName)
    docNames = docName.split(",")
    print("x"*10)
    print(docNames)

    print("x"*10)
    numFiles = 0
    for index, file in enumerate(docNames):
        print(docNames[index])
        if file != "":
            numFiles +=1

    print(numFiles)

    locOne = docNames[0]

    # once file names are received, calls convert() function
    converter(numFiles, docNames)

# function below receives file names
def requestFileNames(numFiles):

    docNames = []

    num = 0
    while num < numFiles:
        docName = input("Enter file path: ")
        docNames.append(docName)
        num += 1

    # once file names are received, calls convert() function
    converter(numFiles, docNames)

# function below receives number of files
def receiveNumFiles():

    try:
        numFiles = int(input("how many files would you like to convert? "))
    except ValueError:
        print('please only enter number')
        receiveNumFiles()

    requestFileNames(numFiles)


# function below receives file names
def bulkUpdater():

    docName = input("Enter file paths (paths must be in same line): ")
    docName = docName.replace(".csv /", ".csv,/")
    #print(docName)
    docNames = docName.split(",")
    #print("x"*10)
    #print(docNames)

    #print("x"*10)
    numFiles = 0
    for index, file in enumerate(docNames):
        #print(docNames[index])
        if file != "":
            numFiles +=1

    #print(numFiles)

    locOne = docNames[0]

    # once file names are received, calls convert() function
    converter(numFiles, docNames)

def main():
    reply = input("""How would you like to convert files:
                A) Enter paths one by one
                B) Enter paths in bulk"
                """)

    if reply == "A":
        receiveNumFiles()
    elif reply == "B":
        bulkUpdater()
    else:
        print("Please enter A or B")
main()
