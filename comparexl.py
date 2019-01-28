from openpyxl import load_workbook
from openpyxl import Workbook
import openpyxl

#Global strings for file names
sans = "Report1548267424539.xlsx"
active = "ActiveEE.xlsx"

class ActiveRow:
    """Row class represents 4 Cells from each row in an excel sheet"""

    def __init__(self,email, fname, lname, empnum):
        """Create a ActiveRow containing cells from the given row"""
        self.email = email
        self.fname = fname
        self.lname = lname
        self.empnum = empnum

    def returnEmail(self):
        """Pull the email from a row, returns a string"""
        return self.email

    def returnFName(self):
        """Pull the first name from a row, returns a string"""
        return self.fname

    def returnLName(self):
        """Pull the last name string from a row, returns a string"""
        return self.lname

    def returnEmpNum(self):
        """Pulls the employee number from a row, returns an int"""
        return self.empnum

    def toString(self):
        return "{}".format(self.email)

    def compareRow(self, row):
        """Compare the email string between two rows, returns a Boolean value"""
        self.row = row
        email = self.email
        email2 = row.returnEmail
        if (email == email2):
            return True
        else:
            return False

##SANS Functions##
def loadSANS():
    #Workbook and worksheet to pull data from
    wb = load_workbook(sans, data_only=True)
    ws = wb.active
    #List to store emails from wb
    sansRows = []
    print("Loading data from " + sans)
    for row in range(3,ws.max_row+1):
        formEmail = "{}{}".format("E", row)
        formFName = "{}{}".format("C", row)
        formLName = "{}{}".format("B", row)
        formEmpNum = "{}{}".format("A", row)
        temprows = ActiveRow(ws[formEmail].value, ws[formFName].value, ws[formLName].value, ws[formEmpNum].value)
        #Adding the cell reference.value to list
        sansRows.append(temprows)
    print("Loaded data from Report1548267424539.xlsx\n")
    return sansRows

##ActiveEE Functions##

def loadActive():
    #Loads emails
    #Workbook and worksheet to pull data from
    wb = load_workbook(active, data_only=True)
    ws = wb.active
    activeRows = []
    print("Loading data from " + active)
    for row in range(2,ws.max_row+1):
        formEmail = "{}{}".format("D", row)
        formFName = "{}{}".format("A", row)
        formLName = "{}{}".format("B", row)
        formEmpNum = "{}{}".format("C", row)
        temprows = ActiveRow(ws[formEmail].value, ws[formFName].value, ws[formLName].value, ws[formEmpNum].value)
        #Adding the cell reference.value to list
        activeRows.append(temprows)
    print("Loaded emails from activeEE.xlsx\n")
    return(activeRows)

##Data Functions

def findInactiveEmails():
    #Pulls list data from other functions and compares the data to find
    #Inactive and new users, creating a new list
    print("Searching for inactive users\n")
    listSANS = loadSANS()
    listActive = loadActive()
    convertedSANS = []
    convertedActive = []
    listInactive = []

    #Find inactive users
    #Convert lists of objects to list of emails
    for i in range(len(listSANS)):
        convertedSANS.append(listSANS[i].returnEmail())

    for i in range(len(listActive)):
        convertedActive.append(listActive[i].returnEmail())

    for i in range(len(convertedSANS)):
        try:
            convertedActive.index(convertedSANS[i])
        except ValueError:
            listInactive.append(listSANS[i])

    print("Finished searching\n")
    print(str(len(listInactive)))
    return(listInactive)

def findNewEmails():
    print("Searching for new users\n")
    listSANS = loadSANS()
    listActive = loadActive()
    convertedSANS = []
    convertedActive = []
    listNew = []

    for i in range(len(listSANS)):
        convertedSANS.append(listSANS[i].returnEmail())

    for i in range(len(listActive)):
        convertedActive.append(listActive[i].returnEmail())

    for i in range(len(convertedActive)):
        try:
            convertedSANS.index(convertedActive[i])
        except ValueError:
            listNew.append(listActive[i])

    print("Finished searching\n")
    print(str(len(listNew)))
    return(listNew)

def exportData():
    inactiveEmails = findInactiveEmails()
    newEmails = findNewEmails()

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
    for x in range(len(inactiveEmails)):
        ws.cell(row=i, column = 4).value = inactiveEmails[x].returnEmail()
        ws.cell(row=i, column = 7).value = "NO"
        ws.cell(row=i, column = 8).value = "NO"
        ws.cell(row=i, column = 3).value = inactiveEmails[x].returnFName()
        ws.cell(row=i, column = 2).value = inactiveEmails[x].returnLName()
        ws.cell(row=i, column = 1).value = inactiveEmails[x].returnEmpNum()
        i+=1
    print("Write successful\n")

    print("Attempting to write new users to file")
    for x in range(len(newEmails)):
        ws.cell(row=i, column = 4).value = newEmails[x].returnEmail()
        ws.cell(row=i, column = 7).value = "YES"
        ws.cell(row=i, column = 8).value = "YES"
        ws.cell(row=i, column = 3).value = newEmails[x].returnFName()
        ws.cell(row=i, column = 2).value = newEmails[x].returnLName()
        ws.cell(row=i, column = 1).value = newEmails[x].returnEmpNum()
        i+=1
    print("Write successful")
    print("Saving file...")
    wb.save('updated.xlsx')
    print("updated.xlsx saved")

def main():
    exportData()
main()
