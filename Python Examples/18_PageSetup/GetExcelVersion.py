from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/ExcelSample_N1.xlsx"
outputFile = "GetExcelVersion.txt"

builder = []
#Create a workbook
workbook = Workbook()
#Load the document
workbook.LoadFromFile(inputFile)
#Get the version
version = workbook.Version
builder.append(str(version))
#Save to file
AppendAllText(outputFile, builder)
workbook.Dispose()

