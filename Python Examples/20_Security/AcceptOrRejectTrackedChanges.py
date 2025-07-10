from spire.xls import *
from spire.xls.common import *


inputFile = "./Demos/Data/TrackChanges.xlsx"
outputFile = "AcceptOrRejectTrackedChanges.xlsx"

workbook = Workbook()
workbook.LoadFromFile(inputFile)
#Accept the changes or reject the changes.
#workbook.AcceptAllTrackedChanges()
workbook.RejectAllTrackedChanges()
#Save to file.
workbook.SaveToFile(outputFile, FileFormat.Version2013)
workbook.Dispose()
