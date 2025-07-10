from spire.xls import *
from spire.xls.common import *

def AppendAllText(fname:str,text:List[str]):
    fp = open(fname,"w")
    for s in text:
        fp.write(s + "\n")
    fp.close()

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
AppendAllText(outputFile, sb)
workbook.Dispose()

