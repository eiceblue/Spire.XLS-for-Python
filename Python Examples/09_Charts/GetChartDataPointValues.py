import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartToImage.xlsx"
outputFile = "GetChartDataPointValues.txt"

sb = []
#Load the document from disk
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first sheet
sheet = workbook.Worksheets[0]
#Get the chart
chart = sheet.Charts[0]
#Get the first series of the chart
cs = chart.Series[0]
for cr in cs.Values:
    sb.append(cr.RangeAddress + "\r\n")
    #Get the data point value
    sb.append("The value of the data point is " + cr.Value + "\r\n")
#Save and launch result file
File.AppendAllText(outputFile, sb)
workbook.Dispose()
