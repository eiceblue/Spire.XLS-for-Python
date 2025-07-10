from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/ToSVG.xlsx"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
i=0
for worksheet in workbook.Worksheets:
    fs = Stream("sheet-"+str(i)+".svg")
    worksheet.ToSVGStream(fs, 0, 0, 0, 0)
    fs.Flush()
    fs.Close()
    i=i+1
workbook.Dispose()

