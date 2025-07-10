from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

inputFile = "./Demos/Data/MacroSample.xls"
outputFile = "DetectVBAMacros.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Detect if the Excel file contains VBA macros
value = ""
hasMacros = workbook.HasMacros
if hasMacros:
    value = "Yes"
else:
    value = "No"
AppendAllText(outputFile, value)

