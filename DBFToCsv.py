# Convert the DBF files to excel
import os
import sys
import csv
from dbfread import DBF

# Source dbf files
dbfFilePath = "*** dbf Files Path *****"

# Destination CSV files
csvFilePath = "*** Destination CSV Path ****"

# source files to get dbfname
files = os.listdir(dbfFilePath)
dbfFiles = []

### Filter DBF file names
for f in files:
    name, ext = os.path.splitext(f)
    if ext == ".DBF" or ext == ".dbf":
        print(f.split(".")[0])
        dbfFiles.append(f)
    

### Convert those DBF Files to CSV

for dbfFile in dbfFiles:
    dbfTable = DBF(dbfFilePath + "\\" + dbfFile, ignore_missing_memofile = True)
    print("Opening File: " + csvFilePath + "\\" + dbfFile.split(".")[0] + ".csv")
    with open(csvFilePath + "\\" + dbfFile.split(".")[0] + ".csv", "w") as csvfile:
        print("File opened successfully...\n")
        csvWriter = csv.writer(csvfile)
        print("Writing header values: \n")
        print(dbfTable.field_names)
        csvWriter.writerow(dbfTable.field_names)
        with dbfTable as table:
            for rec in table:
                #print("row: \n")
                #print(rec.values())
                csvWriter.writerow(rec.values())
