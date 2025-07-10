from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

inputFile = "./Demos/Data/SampeB_4.xlsx"
outputFile = "GetCategoryLabels.txt "

sb = []
#Create a workbook
workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Get the cell range of the category labels
cr = chart.PrimaryCategoryAxis.CategoryLabels
for cell in cr:
    sb.append(cell.Value + "\r\n")
#Save and launch result file  
AppendAllText(outputFile, sb)
workbook.Dispose()
