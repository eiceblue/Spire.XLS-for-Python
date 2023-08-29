from spire.xls import *
from spire.common import *


inputFile = "./Demos/Data/ToHtml.xlsx"
outputFile = "ToHtml.html"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
sheet = workbook.Worksheets[0]
options = HTMLOptions()
options.ImageEmbedded = True
sheet.SaveToHtml(outputFile)
workbook.Dispose()


