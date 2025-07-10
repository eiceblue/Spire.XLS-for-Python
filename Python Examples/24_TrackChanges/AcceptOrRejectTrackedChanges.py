from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/TrackChanges.xlsx"
outputFile = "AcceptOrRejectTrackedChanges.xlsx"

#create a workbook
workbook = Workbook()
#load an excel document
workbook.LoadFromFile(inputFile)
#accept the changes or reject the changes
#workbook.AcceptAllTrackedChanges()
workbook.RejectAllTrackedChanges()

#save the file
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()


