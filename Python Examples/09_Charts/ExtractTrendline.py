import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartSample4.xlsx"
outputFile = "ExtractTrendline.txt"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the chart from the first worksheet
chart = workbook.Worksheets[0].Charts[0]
#Get the trendline of the chart and then extract the equation of the trendline
trendLine = chart.Series[1].TrendLines[0]
formula = trendLine.Formula
sb = []
sb.append("The equation is: " + formula)
#Save to Text file
File.AppendAllText(outputFile, sb)
workbook.Dispose()

