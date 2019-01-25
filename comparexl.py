from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl

#Global strings for file names
sans = "Report1548267424539.xlsx"
active = "ActiveEE.xlsx"


##SANS Functions##
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
            tempString = str(ws[cell_name].value).strip().lower()
            #Adding the cell reference.value to list
            emails.append(tempString)
    print("Loaded emails from Report1548267424539.xlsx\n")
    return(emails)

def loadSANSFNames():
    wb = load_workbook(sans)
    ws = wb.active
    #return list
    sansFNames = []
    #Temporary string to split first and last names
    toSplit = ""
    print("Parsing names from Report1548267424539.xlsx")
    for row in range(3,ws.max_row+1):
        for column in "C":
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value).strip().lower()
            sansFNames.append(tempString)
    print("Loaded names from Report1548267424539.xlsx")
    return sansFNames

def loadSANSLNames():
    wb = load_workbook(sans)
    ws = wb.active
    sansLNames = []

    for row in range(3, ws.max_row+1):
        for column in "B":
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value).strip().lower()
            sansLNames.append(tempString)
    return sansLNames

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
    return sansEmpNum

##ActiveEE Functions##

def loadActiveFNames():
    wb = load_workbook(active)
    ws = wb.active
    activeFNames = []

    print("Parsing names from ActiveEE.xlsx")
    for row in range(2,ws.max_row+1):
        for column in "A":
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value).strip().lower()
            activeFNames.append(tempString)
    print("Loaded names from ActiveEE.xlsx")
    return activeFNames

def loadActiveLNames():
    wb = load_workbook(active)
    ws = wb.active
    activeLNames = []

    for row in range(2,ws.max_row+1):
        for column in "B":
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value).strip().lower()
            activeLNames.append(tempString)
    return activeLNames

def loadActiveEmpNum():
    #Employee numbers must be 9 characters long with 0's added at the beginning
    wb = load_workbook(active)
    ws = wb.active
    tempString =""
    activeEmpNums = []

    for row in range(2, ws.max_row+1):
        for column in "C":
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value)
            tempString.zfill(9)
            activeEmpNums.append(tempString)
    return activeEmpNums

def loadActive():
    #Loads emails
    #Workbook and worksheet to pull data from
    wb = load_workbook(active, data_only=True)
    ws = wb.active
    tempString = ""
    #List to store emails from wb
    activeEmails = []

    #Looping through column B in ActiveEE.xlsx
    print("Processing data from activeEE.xlsx...")
    for row in range(2,ws.max_row+1):
        for column in "D":
            #Grabbing cell reference
            cell_name = "{}{}".format(column, row)
            tempString = str(ws[cell_name].value).strip().lower()
            #Adding the cell reference.value list
            activeEmails.append(tempString)
    print("Loaded emails from activeEE.xlsx\n")
    return(activeEmails)

##Data Functions
def findNewFNames():
    sansFNames = loadSANSFNames()
    activeFNames = loadActiveFNames()

    newActiveFNames = []

    # Find inactive users
    for i in range(len(activeFNames)):
        try:
            sansFNames.index(activeFNames[i])
        except ValueError:
            newActiveFNames.append(activeFNames[i])

    print("Finished searching for active users names\n")
    return (newActiveFNames)

def findNewLNames():
    sansLNames = loadSANSLNames()
    activeLNames = loadActiveLNames()
    newActiveLNames = []

    for i in range(len(activeLNames)):
        try:
            sansLNames.index(activeLNames[i])
        except ValueError:
            newActiveLNames.append(activeLNames[i])
    return newActiveLNames

def findInactiveFNames():
    sansFNames = loadSANSFNames()
    activeFNames = loadActiveFNames()
    newInactiveFNames = []

    for i in range(len(sansFNames)):
        try:
            activeFNames.index(sansFNames[i])
        except ValueError:
            newInactiveFNames.append(sansFNames[i])
    return newInactiveFNames

def findInactiveLNames():
    sansLNames = loadSANSLNames()
    activeLNames = loadActiveLNames()
    newInactiveLNames = []

    for i in range(len(sansLNames)):
        try:
            activeLNames.index(sansLNames[i])
        except ValueError:
            newInactiveLNames.append(sansLNames[i])
    return newInactiveLNames

def findNewEmpNum():
    sansEmpNum = loadSANSEmployeeNums()
    activeEmpNum = loadActiveEmpNum()
    newEmpNum = []

    for i in range(len(activeEmpNum)):
        try:
            sansEmpNum.index(activeEmpNum[i])
        except ValueError:
            newEmpNum.append(activeEmpNum[i])
    return newEmpNum

def findInactiveEmpNum():
    sansEmpNum = loadSANSEmployeeNums()
    activeEmpNum = loadActiveEmpNum()
    inactiveEmpNum = []

    for i in range(len(sansEmpNum)):
        try:
            activeEmpNum.index(sansEmpNum[i])
        except ValueError:
            inactiveEmpNum.append(sansEmpNum[i])
    return inactiveEmpNum

def findInactiveEmails():
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

def findNewEmails():
    print("Searching for new users\n")
    listSANS = loadSANS()
    listActive = loadActive()
    listNew = []

    for i in range(len(listActive)):
        try:
            listSANS.index(listActive[i])
        except ValueError:
            listNew.append(listActive[i])

    print("Finished searching\n")
    return(listNew)

def exportData():
    listInactiveEmails = findInactiveEmails()
    listNewEmails = findNewEmails()
    inactiveFNames = findInactiveFNames()
    inactiveLNames = findInactiveLNames()
    newFNames = findNewFNames()
    newLNames = findNewLNames()
    newEmpNums = findNewEmpNum()
    inactiveEmpNums = findInactiveEmpNum()

    print(str(len(listInactiveEmails)) + "\n" + str(len(listNewEmails)) + "\n" + str(len(inactiveFNames)) +"\n"+ str(len(inactiveLNames)) +"\n"+ str(len(newFNames)) +"\n"+ str(len(newLNames)) +"\n"+ str(len(newEmpNums)) +"\n"+ str(len(inactiveEmpNums)))

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
    for x in range(len(listInactiveEmails)):
        ws.cell(row=i, column = 4).value = listInactiveEmails[x]
        ws.cell(row=i, column = 7).value = "NO"
        ws.cell(row=i, column = 8).value = "NO"
        ws.cell(row=i, column = 3).value = inactiveFNames[x]
        ws.cell(row=i, column = 2).value = inactiveLNames[x]
        ws.cell(row=i, column = 1).value = inactiveEmpNums[x]
        i+=1
    print("Write succesful\n")

    print("Attempting to write new users to file")
    for x in range(len(listNewEmails)):
        ws.cell(row=i, column = 4).value = listNewEmails[x]
        ws.cell(row=i, column = 7).value = "YES"
        ws.cell(row=i, column = 8).value = "YES"
        ws.cell(row=i, column = 3).value = newFNames[x]
        ws.cell(row=i, column = 2).value = newLNames[x]
        ws.cell(row=i, column = 1).value = newEmpNums[x]
        i+=1

    print("Saving file...")
    wb.save('updated.xlsx')
    print("updated.xlsx saved")
def main():
    exportData()
main()
