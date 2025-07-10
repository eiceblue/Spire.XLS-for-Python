from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/CheckAutoFitRowsAndColumns.xlsx"
outputFile = "CheckAutoFitRowOrColumn.txt"

#Create a workbook
workbook = Workbook()
result = []
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Check whether the cell has an adaptive row height set
isRowAutofit = sheet.GetRowIsAutoFit(2)
if isRowAutofit:
        result.append("The second row is auto fit row height.")
else:
        result.append("The second row is not auto fit row height.")
#Check whether the cell has an adaptive column width set
isColAutofit = sheet.GetColumnIsAutoFit(2)
if isColAutofit:
        result.append("The second column is auto fit column width.")
else:
        result.append("The second column is not auto fit column width.")
AppendAllText(outputFile, result)
workbook.Dispose()