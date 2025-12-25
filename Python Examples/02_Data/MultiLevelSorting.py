from spire.xls import *

inputFile = "./Demos/Data/MultiLevelSorting.xlsx"
outputFile = "Sorting_output.xlsx"

#Create a workbook
workbook = Workbook()
#Load the document from disk
workbook.LoadFromFile(inputFile)

# Loop through all worksheets in the workbook
for i in range(0, workbook.Worksheets.Count):
    # Get current worksheet by index
    worksheet = workbook.Worksheets.get_Item(i)

    # Sheet 0: Two-level sort based on cell values
    if i == 0:
        # Clear previous sort rules
        workbook.DataSorter.SortColumns.Clear()
        # Add first sort column: column 0 (A), sort by values in ascending order
        workbook.DataSorter.SortColumns.Add(0, SortComparsionType.Values, OrderBy.Ascending)
        # Add second sort column: column 1 (B), sort by values in descending order
        workbook.DataSorter.SortColumns.Add(1, SortComparsionType.Values, OrderBy.Descending)
        # Apply sorting to data
        workbook.DataSorter.Sort(worksheet.AllocatedRange)

    # Sheet 1: Sort by background color (move dark colors to bottom in column B, light colors to top in column C)
    elif i == 1:
        workbook.DataSorter.SortColumns.Clear()
        # Sort column 1 (B) by background color, dark colors to bottom
        workbook.DataSorter.SortColumns.Add(1, SortComparsionType.BackgroundColor, OrderBy.Bottom)
        # Sort column 2 (C) by background color, light colors to top
        workbook.DataSorter.SortColumns.Add(2, SortComparsionType.BackgroundColor, OrderBy.Top)
        workbook.DataSorter.Sort(worksheet.AllocatedRange)

    # Sheet 2: Sort by font color (light font to top in column B, dark font to bottom in column D)
    elif i == 2:
        workbook.DataSorter.SortColumns.Clear()
        # Sort column 1 (B) by font color, light fonts to top
        workbook.DataSorter.SortColumns.Add(1, SortComparsionType.FontColor, OrderBy.Top)
        # Sort column 3 (D) by font color, dark fonts to bottom
        workbook.DataSorter.SortColumns.Add(3, SortComparsionType.FontColor, OrderBy.Bottom)
        workbook.DataSorter.Sort(worksheet.AllocatedRange)

    # Sheet 3: Sort by conditional formatting icons (high icons to top in column F, low icons to bottom in column G)
    elif i == 3:
        workbook.DataSorter.SortColumns.Clear()
        # Sort column 5 (F) by icon, high icons to top
        workbook.DataSorter.SortColumns.Add(5, SortComparsionType.Icon, OrderBy.Top)
        # Sort column 6 (G) by icon, low icons to bottom
        workbook.DataSorter.SortColumns.Add(6, SortComparsionType.Icon, OrderBy.Bottom)
        workbook.DataSorter.Sort(worksheet.AllocatedRange)

    # Sheet 4: Multi-level sort (Value → BackgroundColor → FontColor → Icon)
    elif i == 4:
        workbook.DataSorter.SortColumns.Clear()
        # First level: sort column 0 (A) by values in ascending order
        workbook.DataSorter.SortColumns.Add(0, SortComparsionType.Values, OrderBy.Ascending)
        # Second level: sort column 2 (C) by background color, dark to bottom
        workbook.DataSorter.SortColumns.Add(2, SortComparsionType.BackgroundColor, OrderBy.Bottom)
        # Third level: sort column 3 (D) by font color, dark to bottom
        workbook.DataSorter.SortColumns.Add(3, SortComparsionType.FontColor, OrderBy.Bottom)
        # Fourth level: sort column 5 (F) by icon, high icons to top
        workbook.DataSorter.SortColumns.Add(5, SortComparsionType.Icon, OrderBy.Top)
        workbook.DataSorter.Sort(worksheet.AllocatedRange)

# Save the updated workbook to a new file with Excel 2016 format
workbook.SaveToFile(outputFile, ExcelVersion.Version2016)
# Release resources and close the workbook
workbook.Dispose()