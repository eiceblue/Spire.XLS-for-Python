from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

outputFile = "GetDefaultRowAndColumnCount.txt"

#Create a workbook
workbook = Workbook()
#Clear all worksheets
workbook.Worksheets.Clear()
#Create a new worksheet
sheet = workbook.CreateEmptySheet()
sb = []
#Get row and column count
rowCount = sheet.Rows.Length
columnCount = sheet.Columns.Length
sb.append("The default row count is :" + str(rowCount))
sb.append("The default column count is :" + str(columnCount))
AppendAllText(outputFile, sb)

