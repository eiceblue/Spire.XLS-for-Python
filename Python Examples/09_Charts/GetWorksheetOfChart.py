import os
import sys
curPath = os.path.abspath(os.path.dirname(__file__))
rootPath = os.path.split(curPath)[0]
sys.path.append(rootPath)
from TestUtil.File import *
from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartToImage.xlsx"
outputFile = "GetWorksheetOfChart.txt"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Access first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Access the first chart inside this worksheet
chart = worksheet.Charts[0]
#Get its worksheet
obj = chart.Worksheet
wSheet = Worksheet(obj)
#Create StringBuilder to save 
content = []
#Set string format for displaying
result = "Sheet Name: " + worksheet.Name + "\r\nCharts' sheet Name: " + wSheet.Name
#Add result string to StringBuilder
content.append(result)
#Save them to a txt file
File.AppendAllText(outputFile, content)
workbook.Dispose()

