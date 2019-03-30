import os
from dbfread import DBF
import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
import time
import shutil
import logging


logging.basicConfig(level=logging.INFO, filename="dbfToExcel.log",
                    format="%(asctime)s | %(levelname)s | %(message)s")
sourceFolder = os.getcwd()  # .py file path


# os.path.abspath(os.path.dirname(__file__)) # .py file path: the same.
# dbfFileList = []
startTime = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(time.time()))
print("Start program at: ", startTime)
folderDict = {}
# folderList = []
# xlsFileList = []

def createExcel(xlsName):
    xlsWriter = pd.ExcelWriter(xlsName, engine='xlsxwriter')
    xlsWriter.save()
    xlsWriter.close()

def checkXlsFile(excelFileFolder, xlsFile, folderPath, folder):
    if not os.path.exists(excelFileFolder + "/archive"):
        os.mkdir(excelFileFolder + "/archive")
    if os.path.exists(xlsFile):
        actualDate = time.strftime('%Y-%m-%d', time.localtime(os.path.getctime(xlsFile)))
        logging.info("{} is exist.".format(xlsFile))
        logging.info("Move {} to archive folder".format(xlsFile))
        shutil.move(xlsFile.replace("\\", "/"), excelFileFolder.replace("\\", "/") + "/archive" + "/" + folder + "_" + actualDate + ".xlsx")
    else:
        pass

def checkSavingFolder(excelFileFolder):
    if not os.path.exists(excelFileFolder):
        logging.info("Create {}".format(excelFileFolder))
        os.mkdir(excelFileFolder)
    else:
        pass

def createFolderList(folderPath):
    excelFileFolder = folderPath + "\\" + "dbfToExcel"
    checkSavingFolder(excelFileFolder)
    for folder in os.listdir(folderPath):
        if os.path.isdir(folder) and folder !=".idea" and  folder !="dbfToExcel":
            # global folderList
            # folderList.append(folderPath + "\\" + folder)
            dbfFileList = createDbfList(folderPath + "\\" + folder)
            xlsFile = excelFileFolder + "\\" + folder + ".xlsx"
            global folderDict
            folderDict[folder] = (folderPath + "\\" + folder, xlsFile, dbfFileList)
            logging.info("Create Dictionary from {} folder".format(folder))
            # logging.info("SUM:  {} + {} = {}".format(num1, num2, result))
            logging.info("Check Excels")
            checkXlsFile(excelFileFolder, xlsFile, folderPath,  folder)
            createExcel(xlsFile)
            # xlsFileList.append(xlsFile)
        else:
            continue

def createDbfList(folder):
    logging.info("Get dbf list in {} ".format(folder))
    dbfFileList = [i for i in os.listdir(folder) if i.upper().endswith(".DBF")]
    dbfFileDict = {}
    dbfFileDict["Dbfile"] = dbfFileList
    return dbfFileDict
    logging.info("Create DBF file List: {}".format(dbfFileList))

def saveDbf(folder, xlsFile,dbfList):
    logging.info("Save {} into {}".format(dbfList,xlsFile))
    book = load_workbook(xlsFile)
    writer = pd.ExcelWriter(xlsFile, engine='openpyxl')
    logging.info("Open {}".format(xlsFile))
    writer.book = book
    print("Open ", xlsFile)
    for i in dbfList:
        logging.info("Create {} dataframe and save into sheet".format(i))
        dbf = DBF(folder + "\\" + i)
        frame = pd.DataFrame(iter(dbf))
        frame.to_excel(writer, sheet_name=i)
    logging.info("Save {} file".format(xlsFile))
    print("Save {}".format(xlsFile))
    writer.save()
    logging.info("Close {} ".format(xlsFile))
    writer.close()
    logging.info(50 * "--")

def main():
    createFolderList(sourceFolder)
    for key,value in folderDict.items():
        logging.info("Start to saving DBF file into: {}".format(value[1]))
        saveDbf(value[0], value[1],  value[2]['Dbfile'])

main()
endTime = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime(time.time()))
print("Saved files at: ", endTime)
logging.info("StartTime: {}".format(startTime))
logging.info("EndTime: {}".format(endTime))
#input("Press enter to exit ;)")