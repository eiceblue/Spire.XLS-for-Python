from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

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

