from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/PageBreak.xlsx"
outputFile = "RemovePageBreak.xlsx"

#Create a workbook and load a file
workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Get the first worksheet from the workbook
sheet = workbook.Worksheets[0]
#Clear all the vertical page breaks
sheet.VPageBreaks.Clear()
#Remove the firt horizontal Page Break
sheet.HPageBreaks.RemoveAt(0)
#Set the ViewMode as Preview to see how the page breaks work
sheet.ViewMode = ViewMode.Preview
#Save the document
workbook.SaveToFile(outputFile, ExcelVersion.Version2013)
workbook.Dispose()
