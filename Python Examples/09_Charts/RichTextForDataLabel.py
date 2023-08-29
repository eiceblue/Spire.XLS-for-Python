from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ChartToImage.xlsx"
outputFile = "RichTextForDataLabel.xlsx"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get first worksheet of the workbook
worksheet = workbook.Worksheets[0]
#Get the first chart inside this worksheet
chart = worksheet.Charts[0]
#Get the first datalabel of the first series 
datalabel = chart.Series[0].DataPoints[0].DataLabels
#Set the text
datalabel.Text = "Rich Text Label"
#Show the value
chart.Series[0].DataPoints[0].DataLabels.HasValue = True
#Set styles for the text
#chart.Series[0].DataPoints[0].DataLabels.Font.Color = Color.get_Red()
#chart.Series[0].DataPoints[0].DataLabels.Font.IsBold = true
chart.Series[0].DataPoints[0].DataLabels.Color = Color.get_Red()
chart.Series[0].DataPoints[0].DataLabels.IsBold = True
#Save the file
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
