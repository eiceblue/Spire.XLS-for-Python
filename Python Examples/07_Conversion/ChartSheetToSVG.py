from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ChartSheet.xlsx"
outputFile = "ChartSheetToSVG.svg"
       
#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)
#Get the second chartsheet by name
cs = workbook.GetChartSheetByName("Chart1")
fs = Stream(outputFile)
cs.ToSVGStream(fs)
fs.Flush()
fs.Close()

