from spire.xls import *
from spire.xls.common import *

def AppendText(fname:str,text:str):
    fp = open(fname,"w")
    fp.write(text + "\n")
    fp.close()

inputFile = "./Demos/Data/ProtectedWorkbook.xlsx"
outputFile = "DetectProtection.txt"

value = Workbook.IsPasswordProtected(inputFile)
boolvalue = ""
if value:
    boolvalue = "Yes"
else:
    boolvalue = "No"
AppendText(outputFile, boolvalue)

