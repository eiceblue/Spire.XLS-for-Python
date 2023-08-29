from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ReadImages.xlsx"
outputFile = "ToHtmlStream.html"

#Create a workbook
workbook = Workbook()
#Load the Excel document from disk
workbook.LoadFromFile(inputFile)
#Get the first worksheet
sheet = workbook.Worksheets[0]
#Set the html options
options = HTMLOptions()
options.ImageEmbedded = True
#Save sheet to html stream
fileStream = Stream(outputFile)
sheet.SaveToHtml(fileStream, options)
fileStream.Close()
workbook.Dispose()
