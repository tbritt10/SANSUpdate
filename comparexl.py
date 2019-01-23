from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl

#Global strings for file names
sans = "Report1548267424539.xlsx"
active = "ActiveEE.xlsx"

def loadSANS():
    #Workbook and worksheet to pull data from
        wb = load_workbook(sans, data_only=True)
        ws = wb.active
        #List to store emails from wb
        emails = []
        print("Processing data from Report1548267424539.xlsx")
        for row in range(3,ws.max_row+1):
            for column in "E":
                #Grabbing cell reference
                cell_name = "{}{}".format(column, row)
                #Adding the cell reference.value to list
                emails.append(ws[cell_name].value)
        print("Loaded emails from Report1548267424539.xlsx\n")
        return(emails)

def loadSANSNames():
    wb = load_workbook(sans)
    ws = wb.active
    #return list
    sansNames = []
    #Temporary string to split first and last names
    toSplit = ""
    print("Parsing names from Report1548267424539.xlsx")
    for row in range(3,ws.max_row+1):
        for column in "B":
            cell_name = "{}{}".format(column, row)
            toSplit = ws[cell_name].value
            sansNames.append(toSplit.split())
    print("Loaded names from Report1548267424539.xlsx")
    return sansNames
def loadSANSEmployeeNums():
    wb = load_workbook(sans)
    ws = wb.active
    #return list
    sansEmpNum = []
    print("Parsing employee numbers from Report1548267424539.xlsx")
    for row in range(3,ws.max_row+1):
        for column in "A":
            cell_name = "{}{}".format(column, row)
            sansEmpNum.append(ws[cell_name].value)
    print("Succesfully loaded employee numbers from Report1548267424539.xlsx")

def findInactiveNames():
    sansNames = loadSANSNames()
    activeNames = loadActiveNames()

    inactiveNames = []

    # Find inactive users
    for i in range(len(sansNames)):
        try:
            activeNames.index(sansNames[i])
        except ValueError:
            inactiveNames.append(sansNames[i])

    print("Finished searching for inactive users names\n")
    return (inactiveNames)

def findActiveNames():
    sansNames = loadSANSNames()
    activeNames = loadActiveNames()

    newActiveNames = []

    # Find inactive users
    for i in range(len(activeNames)):
        try:
            sansNames.index(activeNames[i])
        except ValueError:
            newActiveNames.append(activeNames[i])

    print("Finished searching for active users names\n")
    return (newActiveNames)

def loadActiveNames():
    wb = load_workbook(active)
    ws = wb.active
    #return list
    activeNames = []
    #Temporary string to split first and last names
    toSplit = ""
    print("Parsing names from ActiveEE.xlsx")
    for row in range(2,ws.max_row+1):
        for column in "A":
            cell_name = "{}{}".format(column, row)
            toSplit = ws[cell_name].value
            #Need to strip comma and one space - not currently working
            toSplit.strip(' ,')
            activeNames.append(toSplit.split())
    print("Loaded names from ActiveEE.xlsx")
    for name in range(len(activeNames)):
        print(activeNames[name])
    return activeNames

def loadActive():

        #Workbook and worksheet to pull data from
        wb = load_workbook(active, data_only=True)
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
    #Inactive and new users, creating a new list
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
        ws.cell(row=i, column = 7).value = "YES"
        ws.cell(row=i, column = 8).value = "YES"
        i+=1

    print("Writing names to file...")
    #Pull names; Compare the lists and create 2 new lists (still nested) from the two. Use .zip(*) to seperate first and last names then write to file.
    #for i in range(len(listActive)):
        #try:
            #listSANS.index(listActive[i])
        #except ValueError:
            #listNew.append(listActive[i])
    #for x in range(len(inactiveFNames)):
        #ws.cell(row=j, column = 3).value = inactiveFNames[x]
        #ws.cell(row=j, column = 2).value = inactiveLNames[x]
    #for x in range(len(activeFNames)):
        #ws.cell(row=j, column = 3).value = activeFNames[x]
        #ws.cell(row=j, column = 2).value = activeLNames[x]

    #Write employee numbers to file
    print("Writing employee numbers to file")
    #sansNums = loadSANSEmployeeNums()
    #activeNums = loadActiveEmployeeNums()

    #for i in range(len(listActive)):
        #try:
            #listSANS.index(listActive[i])
        #except ValueError:
            #listNew.append(listActive[i])

    print("Saving file...")
    wb.save('updated.xlsx')
    print("updated.xlsx saved")
def main():
    exportData()
    #loadSANSNames()
    #loadActiveNames()
main()
