from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()
	

inputFile = "./Demos/Data/WorksheetSample2.xlsx"
outputFile = "GetPaperSize.txt"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#pageInfoList = workbook.GetSplitPageInfo()
sb = []
for sheet in workbook.Worksheets:
    width = sheet.PageSetup.PageWidth
    height = sheet.PageSetup.PageHeight
    sb.append(sheet.Name)
    sb.append("Width: " + str(width) + "\tHeight: " + str(height))
#Save the documen
AppendAllText(outputFile, sb)
workbook.Dispose()

