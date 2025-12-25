from spire.xls.common import *
from spire.xls import *

inputFile = "Data/ProhibitDtd.xlsx"

try:
    # Create a workbook.
    workbook = Workbook()

    # Prohibit DTD
    workbook.ProhibitDtd=True

    # Load the Excel document from disk
    workbook.LoadFromFile(inputFile)

    # Dispose of the workbook object to release resources.
    workbook.Dispose()
except Exception as ex:
    print("The DTD was disabled.")
    print(str(ex))