import xlrd
import xlsxwriter
from xlrd import open_workbook
import os
import os.path
from pathlib import Path
from pathlib import Path
import datetime

def notExistsIn(serialNumber, bigDataSet):
    results = True
    for i in range (0,len(bigDataSet)):
        if(serialNumber in bigDataSet[i]):
            return False
        else:
           results = True
    return results

def getWorkingDirectory(filepath):
    a = str(filepath).split("/")
    b=""
    for i in range(0, len(a)-1):
        b+=a[i]+"/"
    return b

print("Welcome to the Database Upadter!!\n(Developer: Achuna Ofonedu)\n")

trueDatabaseLocation = input("Enter 'True Database' File Location: ").replace("\\", "/").replace("\"","")
#trueDatabaseLocation = ("C:/Users/Achuna/Desktop/Computer Inventory/True Database of Computers.xlsx")
outputDirectory = getWorkingDirectory(trueDatabaseLocation)

trueWb = xlrd.open_workbook(trueDatabaseLocation)
trueSheet = trueWb.sheet_by_index(0)
trueRows = trueSheet.nrows
trueColumns = trueSheet.ncols

trueSerialNumbers = [] #contains all serial numbers in true database
duplicateSerials = [] #rows of duplicates


for i in range(1, trueRows):
    if(trueSheet.cell_value(i, 5) != ""):
        computerSerial = str(trueSheet.cell_value(i, 5)).strip().upper()
        trueSerialNumbers.append(computerSerial)

def checkDups():
    for i in range(0,len(trueSerialNumbers)):
        for j in range(i+1, len(trueSerialNumbers)):
            if(trueSerialNumbers[i] == trueSerialNumbers[j] and (len(str(trueSerialNumbers[j])) > 2 and len(str(trueSerialNumbers[j])) < 10)):
                #print("DUPS: " + trueSerialNumbers[i])
                duplicateSerials.append(trueSerialNumbers[i])
                duplicateSerials.append(trueSerialNumbers[j])

checkDups()
# for i in range(0, len(duplicateSerials)):
#     print(duplicateSerials[i])
inventoryLocation = input("Enter 'Inventory Updates' File Location: ").replace("\\", "/").replace("\"","")
#inventoryLocation = ("C:/Users/Achuna/Desktop/Computer Inventory/NEW Inventory Update (Responses).xlsx")

wb = xlrd.open_workbook(inventoryLocation);
sheet = wb.sheet_by_index(0)
rows = sheet.nrows
columns = sheet.ncols

serialIndex = [15, 18]


allInventory = [] #contains all in new inventory
newInventory = [] #contains computers needed to be updated

def saveRow(sIndex, i): #i is the row
    if sheet.cell(i, 0).ctype == 3:  # 3 means 'xldate' , 1 means 'text'
        ms_date_number = sheet.cell_value(i, 0)  # Correct option 1
        dates = xlrd.xldate_as_tuple(ms_date_number, wb.datemode)
        dateString = str(dates[1]) + "/" + str(dates[2]) + "/" + str(dates[0]) + " " + str(dates[3]) + ":" + str(dates[4]) + ":" + str(dates[5])
        timestamp = dateString
    email = str(sheet.cell_value(i, 1))
    ticket = str(sheet.cell_value(i, 2))
    newCPUName = str(sheet.cell_value(i, 3)).strip().upper()
    oldName = str(sheet.cell_value(i, 4)).strip().upper()
    building = str(sheet.cell_value(i, 5)).strip().upper()
    roomNumber = str(sheet.cell_value(i, 6)).strip()
    location = building + " " + roomNumber
    owner = str(sheet.cell_value(i, 7)).strip()
    department = str(sheet.cell_value(i, 8)).strip()
    modelName = str(sheet.cell_value(i, 9)).strip().upper()
    notused = ""
    partNumber = str(sheet.cell_value(i, 13)).strip()
    otherModelPartNumber = str(sheet.cell_value(i, 14)).strip()
    computerSerial = str(sheet.cell_value(i, serialIndex[sIndex])).strip().upper()
    appleModel = str(sheet.cell_value(i, 16)).strip()
    applePartNumber = str(sheet.cell_value(i, 17)).strip()
    appleSerialNumber = str(sheet.cell_value(i, serialIndex[sIndex])).strip().upper()
    otherManufacturers = str(sheet.cell_value(i, 19)).strip().upper()
    otherMFRSerialNumbers = str(sheet.cell_value(i, 20)).strip().upper()
    purchaseOrder = str(sheet.cell_value(i, 21)).strip()
    OS = str(sheet.cell_value(i, 22)).strip()
    MACAddress = str(sheet.cell_value(i, 23)).upper()
    softwareInstalled = str(sheet.cell_value(i, 24))
    otherSoftware = str(sheet.cell_value(i, 25))
    useOfPC = str(sheet.cell_value(i, 26))
    specialNoters = str(sheet.cell_value(i, 27))
    assetTag = str(sheet.cell_value(i, 28))
    warranty = str(sheet.cell_value(i, 29))



    allInventory.append((timestamp, email, ticket, newCPUName, oldName, building, roomNumber, owner, department,
                         modelName, notused, notused, notused, partNumber, otherModelPartNumber, computerSerial, appleModel,
                         applePartNumber, appleSerialNumber, otherManufacturers, otherMFRSerialNumbers, purchaseOrder, OS, MACAddress,
                         softwareInstalled, otherSoftware, useOfPC, specialNoters, assetTag, warranty))

    if (notExistsIn(computerSerial, trueSerialNumbers)):  ##Time to update true inventory with new element
        newInventory.append((newCPUName, building, department, owner, location, computerSerial, modelName, otherManufacturers, partNumber, warranty, OS, assetTag, notused, purchaseOrder, ticket, MACAddress, softwareInstalled))

for i in range(1,rows):
    if(sheet.cell_value(i, serialIndex[0]) != ""):
       saveRow(0, i)
    elif(sheet.cell_value(i, serialIndex[1]) != ""):
        saveRow(1, i)
    else:
        saveRow(0, i)


#WRITE OUTPUT OF NEW INVENTORY

outputFile = Path(str(outputDirectory)+"/"+"NEW Inventory Updates.xlsx")
if outputFile.is_file():
    os.remove(outputFile)

workbook = xlsxwriter.Workbook(outputFile)
outputSheet = workbook.add_worksheet("Responses")
style1 = workbook.add_format({'font_color' : 'red'})


#Write to a New Inventory Updates file and highlight rows not in True database Red
bold =workbook.add_format({'bold' : True})
for j in range(0, len(allInventory[0])):
    outputSheet.write(0, j, str(sheet.cell_value(0, j)), bold)

for i in range(0,len(allInventory)):
    if (notExistsIn(allInventory[i][15], trueSerialNumbers)):
        for j in range(0, len(allInventory[0])):
                outputSheet.write(i+1, j, str(allInventory[i][j]), style1)
    else:
        for j in range(0, len(allInventory[0])):
                outputSheet.write(i+1, j, str(allInventory[i][j]))

workbook.close()




outputFile2 = Path(str(outputDirectory)+"/"+"NEW True Database.xlsx")
if outputFile2.is_file():
    os.remove(outputFile2)

workbook2 = xlsxwriter.Workbook(outputFile2)
styleD = workbook2.add_format({'bg_color' : 'red'})
outputSheet2 = workbook2.add_worksheet("All_Bldgs")
for i in range(0, trueRows):
    for j in range(0, trueColumns):
        if(j == 5):
            computerSerial = str(trueSheet.cell_value(i, j))
            if (computerSerial in duplicateSerials):
                outputSheet2.write(i, j, str(trueSheet.cell_value(i, j)), styleD)
            else:
                outputSheet2.write(i, j, str(trueSheet.cell_value(i, j)))
        else:
            outputSheet2.write(i, j, str(trueSheet.cell_value(i, j)))


style2 = workbook2.add_format({'font_color' : 'blue'})

start = trueRows
for i in range(0, len(newInventory)):
    for j in range(0, len(newInventory[i])):
        outputSheet2.write(start, j, str(newInventory[i][j]),style2)
    start+=1

workbook2.close()

print("\nSuccessfully Created Output Files!\nOpen (" + str(outputDirectory)+") to view")

input('Press ENTER to exit')