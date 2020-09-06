import openpyxl
import datetime
import time

#==============================================================================
databaseFileName = 'database.xlsx'
exportFileName = 'export.xlsx'

databaseFile = openpyxl.load_workbook(databaseFileName)
exportFile = openpyxl.load_workbook(exportFileName)

# grab the active worksheet
databaseFileSheet = databaseFile.active
exportFileSheet = exportFile.active
#==============================================================================

def countCell():
    cellcount = 3                   #Starting at A3
    cellname = "A" + str(cellcount)
    
    cellVal = databaseFileSheet[cellname].value
    


    while cellVal is not None:
        cellcount += 1
        
        cellname = "A" + str(cellcount)
        cellVal = databaseFileSheet[cellname].value

    return cellcount - 1

#------------------------------------------------------------------------------

namecount = countCell()
#print(namecount)

#------------------------------------------------------------------------------

def checkEmpty(cell):

    if databaseFileSheet[cell].value is None:
        return True
    else:
        return False
    
def formatTimeCell(i):

    day = getDay()
    order = int(ord("C"))
    cell = ""
    
    if day == "Mon":
        temp = str(chr(order)) + str(i)
        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)
        
    elif day == "Tue":
        order += 2
        temp = str(chr(order)) + str(i)

        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)
    elif day == "Wed":
        order += 4
        temp = str(chr(order)) + str(i)

        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)
            
    elif day == "Thu":
        order += 6
        temp = str(chr(order)) + str(i)

        
        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)

        
    elif day == "Fri":
        order += 8
        temp = str(chr(order))+ str(i)

        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)
    else:  #sat or sun code edit
        order += 8
        temp = str(chr(order))+ str(i)

        if checkEmpty(temp):
            cell = temp
        else:
            order += 1
            cell = str(chr(order)) + str(i)

    return str(cell)
    
def setTime():
    time = "15:50"

    return time

def getRealTime():

    dateTime = str(datetime.datetime.now())
    currTime = dateTime[11:16]

    return currTime

def formatTime(currTime):
    
    #if int(currTime[:2]) > 12:
        #loginTime = str(int(currTime[:2])-12) + currTime[2:] + " PM"
    #elif int(currTime[:2]) == 12:
        #loginTime = currTime + " PM"
    #elif int(currTime[:2]) == 0:
        #loginTime = "12" + currTime[2:] + " AM"
    #else:
        #loginTime = currTime + " AM"
    
    loginTime = currTime
    return loginTime

def convertNumTime(loginTime):
    lst = str(loginTime).split(":")

    hour = int(lst[0])
    min = int(str(lst[1])[:2])

    numTime = 100*hour + min
    #print(numTime)

    return numTime

def getDay():
    date = datetime.datetime.now()
    return date.strftime("%a")

def searchName(name):
    i = 0
    for i in range (1,namecount + 1):
    
        cellname = "A" + str(i)
        cellVal = databaseFileSheet[cellname].value
        if name == cellVal:
            pin = str(input("Enter PIN: "))

            checkPIN(pin,i)

def checkPIN(inputPIN,i):

    if inputPIN == str(databaseFileSheet["B" + str(i)].value):
        logTime(i)
        
def logTime(i):
    
    currTime = setTime()
    #currTime = getRealTime()

    loginTime = formatTime(currTime)

    timecell = formatTimeCell(i)
    databaseFileSheet[timecell] = loginTime

    print("You have successfully logged in!")
    print("Name: " + databaseFileSheet["A" + str(i)].value)

    print("TimeCell: " + str(timecell) )
    
    print("Time: " + str(databaseFileSheet[timecell].value))

def evalAttendance(loginCell,logoutCell,currCellExport):

    loginTime = convertNumTime(databaseFileSheet[loginCell].value)
    logoutTime = convertNumTime(databaseFileSheet[logoutCell].value)

    if loginTime <= 730:

        if logoutTime >= 1700:   
            exportFileSheet[currCellExport] = "PRESENT"

        else:
            exportFileSheet[currCellExport] = "INCOMPLETE"
    else:
        if logoutTime >= 1700:
           exportFileSheet[currCellExport] = "LATE"
        else:
            exportFileSheet[currCellExport] = "INCOMPLETE"

    
def exportData():

    for i in range(4,namecount+1): #ranging through all names

        x = 0 #set to 0
        y = 0 #set to 0
        
        while x >= 0 and x < 10: #Column C to L in database (Mon - Fri x2)

            currCellDatabase = str(chr(x+67)) + str(i)
            nextCellDatabase = str(chr(x+68)) + str(i)

            currCellExport = str(chr(y+67)) + str(i)

            #print("Cell = " + currCellExport)
            
            if checkEmpty(currCellDatabase):
                exportFileSheet[currCellExport] = "ABSENT"
            else:
                if not checkEmpty(nextCellDatabase):

                    evalAttendance(currCellDatabase,nextCellDatabase,currCellExport)

                else:  

                    exportFileSheet[currCellExport] = "ABSENT"

            x+=2
            y+=1


def menu():
    print("1 - Login attendance")
    print("2 - Export attendance data")
    print("3 - Exit application")
    
    mode = input("Select an option: ")

    if mode == "1":
        name = str(input("Enter name here: "))
        searchName(name)
    elif mode == "2":
        exportData()
        print("Data exported")
    elif mode == "3":
        print("Exiting application")
        time.sleep(2)
        exit()
        
while True:


    menu()



    databaseFile.save(databaseFileName)
    exportFile.save(exportFileName)

