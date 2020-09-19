# ExcelHelper
C#-class and VBA-module that I created to help working with Excel workbooks.

They both contain the same helper functions, though the VBA-module a little bit more of them. Two that I have found to be the most useful are:

- OpenWb(string path)
- CreateWb()

The functions handle the Excel-application (whether it is running or not) and OpenWb() also whether the requested workbook is already open or not.