from spire.xls import *

inputFile = "./Demos/Data/ToMarkdownExportOptions.xlsx"
outputFile = "ToMarkdownExportOptions_out.md"

# Create a new workbook
workbook = Workbook()

# Load the document
workbook.LoadFromFile(inputFile)

# Create export options for Markdown format
markdownOptions = MarkdownOptions()
# Set whether to save images with relative paths
markdownOptions.SavePicInRelativePath = True
# Set whether to save hyperlinks as Markdown reference format
markdownOptions.SaveHyperlinkAsRef = True

# Save the workbook as Markdown with the specified options
workbook.SaveToMarkdown(outputFile, markdownOptions)
workbook.Dispose()