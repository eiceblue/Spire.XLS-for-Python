from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/RetrieveExternalFileHyperlinks.xlsx"
outputFile = "RetrieveExternalFileHyperlinks.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet.
sheet = workbook.Worksheets[0]
content = []
#Retrieve external file hyperlinks.
for item in sheet.HyperLinks:
    address = item.Address
    sheetName = item.Range.WorksheetName
    range = item.Range
    content.append("Cell[{0},{1}] in sheet \"" + sheetName + "\" contains File URL: {2}".format(range.Row, range.Column, address))
AppendAllText(outputFile, content)
#Save to file
#workbook.SaveToFile(outputFile, ExcelVersion.Version2010)
workbook.Dispose()

