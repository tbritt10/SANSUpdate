from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl

def loadSANS():
    #Workbook and worksheet to pull data from
        wb = load_workbook('SANS.xlsx', data_only=True)
        ws = wb.active
        #List to store emails from wb
        emails = []
        print("Processing data from SANS.xlsx")
        for row in range(2,ws.max_row+1):
            for column in "A":
                #Grabbing cell reference
                cell_name = "{}{}".format(column, row)
                #Adding the cell reference.value to list
                emails.append(ws[cell_name].value)
        print("Loaded emails from SANS.xlsx\n")
        return(emails)

def loadActive():

        #Workbook and worksheet to pull data from
        wb = load_workbook('ActiveEE.xlsx', data_only=True)
        ws = wb.active
        #List to store emails from wb
        activeEmails = []

         #Looping through column B in ActiveEE.xlsx
        print("Processing data from activeEE.xlsx...")
        for row in range(2,ws.max_row+1):  
            for column in "B":  
                #Grabbing cell reference
                cell_name = "{}{}".format(column, row)
                #Adding the cell reference.value list
                activeEmails.append(ws[cell_name].value)
        print("Loaded emails from activeEE.xlsx\n")
        return(activeEmails)

def findInactive():
    #Pulls list data from other functions and compares the data to find
    #Inactive and new users, creating two different lists
    print("Searching for inactive users\n")
    listSANS = loadSANS()
    listActive = loadActive()
    listInactive = []

    #Find inactive users
    for i in range(len(listSANS)):
        try:
            listActive.index(listSANS[i])
        except ValueError:
            listInactive.append(listSANS[i])

    print("Finished searching\n")
    return(listInactive)

def findNew():
    print("Searching for new users\n")
    listSANS = loadSANS()
    listActive = loadActive()
    listInactive = []
    listNew = []

    for i in range(len(listActive)):
        try:
            listSANS.index(listActive[i])
        except ValueError:
            listNew.append(listActive[i])

    print("Finished searching\n")
    return(listNew)

def exportData():
    listInactive = findInactive()
    listNew = findNew()
    i = 2
    print("Attempting to open output file")
    try:
        wb = load_workbook("Employees.xlsx")
        ws = wb.active
        print("Output file loaded succesfully\n")
    except FileNotFoundError:
        print("\nERROR: Employees.xlsx not found")
        print("ERROR: Please download the file from the SANS Administrator portal")
        print("ERROR: The file must be placed in the same directory as the program")
        exit()
    #Put Inactive users into file
    print("Attempting to write inactive users to file")
    for x in range(len(listInactive)):
        ws.cell(row=i, column = 4).value = listInactive[x]
        ws.cell(row=i, column = 7).value = "NO"
        ws.cell(row=i, column = 8).value = "NO"
        i+=1
    print("Write succesful\n")

    print("Attempting to write new users to file")
    for x in range(len(listNew)):
        ws.cell(row=i, column = 4).value = listNew[x]
        i+=1
    print("Saving file...")
    wb.save('updated.xlsx')
    print("updated.xlsx saved")
def main():
    exportData()
main()
