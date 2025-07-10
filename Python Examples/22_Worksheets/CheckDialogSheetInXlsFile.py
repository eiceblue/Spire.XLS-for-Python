from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/CheckDialogSheetInXlsFile.xlsx"
outputFile = "CheckDialogSheetInXlsFile.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
content = []
if sheet.Type == ExcelSheetType.DialogSheet:
    content.append("Worksheet is a Dialog Sheet!")
else:
    content.append("Worksheet is not a Dialog Sheet!")
AppendAllText(outputFile, content)
workbook.Dispose()


