import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


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
File.AppendAllText(outputFile, sb)
workbook.Dispose()
