from spire.xls import *
from spire.xls.common import *


def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/CheckRowOrColumnIsHidden.xlsx"
outputFile = "CheckRowOrColumnIsHidden.txt"

#Create a workbook
workbook = Workbook()
result = []
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Check whether the cell has an adaptive row height set
isRowHide = sheet.GetRowIsHide(2)
if isRowHide:
        result.append("The second row is hidden.")
else:
        result.append("The second row is not hidden.")
#Check whether the cell has an adaptive column width set
isColHide = sheet.GetColumnIsHide(2)
if isColHide:
        result.append("The second column is hidden.")
else:
        result.append("The second column is not hidden.")
AppendAllText(outputFile, result)
workbook.Dispose()