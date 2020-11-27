Attribute VB_Name = "M__Gen_Excel"
Private excelApp As Excel.Application

'Creates and returns a new workbook, regardless of whether there is Excel-application already running.
Function createWb() As Excel.Workbook

    Set excelApp = getExcelAppIfOpen
    
    If TypeName(excelApp) <> "Application" Then
        Set excelApp = CreateObject("Excel.application")
        excelApp.Visible = True
    End If
    
    Set createWb = excelApp.Workbooks.Add
    
End Function

'Opens the workbook specified by the path given as an argument.
'Handles the following situations:
'   - Excel application running / not running
'   - Workbook is open / closed
'   - Another workbook with the same name is open
Function openWb(ByVal path As String) As Excel.Workbook

    Set excelApp = getExcelAppIfOpen
    
    If TypeName(excelApp) <> "Application" Then
        Set excelApp = CreateObject("Excel.application")
        excelApp.Visible = True
        Set openWb = excelApp.Workbooks.Open(path)
    Else
        If isWbOpen(path) Then
            Set openWb = excelApp.Workbooks(nameFromPath(path))
        Else
            If isNamedWbOpen(nameFromPath(path)) Then
                excelApp.Workbooks(nameFromPath(path)).Close True
            End If
            
            Set openWb = excelApp.Workbooks.Open(path)
        End If
    End If

End Function

'Closes the workbook given as an argument. Closes also the Excel-application if there are no more workbooks open.
Sub closeWb(ByRef wb As Excel.Workbook, ByRef whetherToSave As Boolean)

    Set excelApp = GetObject(, "Excel.application")
    
    wb.Close whetherToSave
    
    If excelApp.Workbooks.Count = 0 Then
        excelApp.Quit
    End If

End Sub

'Saves a workbook given as an argument. Handles the situation where there is already a workbook with the given name open.
'Cannot be used to save a workbook that already has a name with that same name. This is acceptable since, in that case, you shouldn't save AS but just save
Sub saveWbAs(ByRef wb As Excel.Workbook, ByVal path As String)

    Set excelApp = GetObject(, "Excel.application")
    
    If isNamedWbOpen(nameFromPath(path)) Then                   'Two workbooks with the same name cannot be open at the same time so the possible other one is closed
        excelApp.Workbooks(nameFromPath(path)).Close True
    End If
    
    excelApp.DisplayAlerts = False                               'So that the user is not prompted for saving
    wb.SaveAs filename:=path
    excelApp.DisplayAlerts = True

End Sub

'Returns the name of the Excel-file that is currently open. If multiple files are open, returns only one name.
Function getOpenExcelName() As String

    Set excelApp = getExcelAppIfOpen
    
    If TypeName(excelApp) = "Application" Then
        If excelApp.Workbooks.Count > 0 Then
            getOpenExcelName = excelApp.Workbooks(1).FullName
        End If
    End If

End Function

Private Function getExcelAppIfOpen() As Excel.Application

    On Error Resume Next
    
    Set getExcelAppIfOpen = GetObject(, "Excel.application")
    
    On Error GoTo 0

End Function

Private Function isWbOpen(ByRef path As String) As Boolean

    Dim wb As Excel.Workbook
    
    For Each wb In excelApp.Workbooks
        If path = wb.FullName Then
            isWbOpen = True
        End If
    Next wb

End Function

Private Function isNamedWbOpen(ByRef name As String) As Boolean

    Dim wb As Excel.Workbook
    
    For Each wb In excelApp.Workbooks
        If name = wb.name Then
            isNamedWbOpen = True
        End If
    Next wb

End Function

Function findLastRow(ByRef sheet As Excel.Worksheet) As Long

    On Error GoTo errorHandler
    
    findLastRow = sheet.Cells.Find("*", After:=sheet.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    Exit Function
errorHandler:
        findLastRow = 0
    
End Function

Function findLastCol(ByRef sheet As Excel.Worksheet) As Integer

    On Error GoTo errorHandler
    
    findLastCol = sheet.Cells.Find("*", After:=sheet.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    Exit Function
errorHandler:
        findLastCol = 0
    
End Function

Function findLastRowByCol(ByRef sheet As Excel.Worksheet, ByVal col1 As String, ByVal col2 As String) As Integer

    On Error GoTo errorHandler
    
    findLastRowByCol = sheet.Range(col1 & ":" & col2).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    Exit Function
errorHandler:
        findLastRowByCol = 0
        
End Function

Function findLastColByRow(ByRef sheet As Excel.Worksheet, ByVal row1 As String, ByVal row2 As String, ByVal asLetter As Boolean) As String

    On Error GoTo errorHandler
    
    Dim colNum As Integer
    colNum = sheet.Range("A" & row1 & ":" & "BA" & row2).Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    If asLetter Then
        findLastColByRow = Split(Cells(1, colNum).Address(True, False), "$")(0)
    Else
        findLastColByRow = CStr(colNum)
    End If
    
    Exit Function
    
errorHandler:
        findLastColByRow = ""
    
End Function

Function selectEverythingOnSheet(ByRef sheet As Excel.Worksheet) As Range

    Dim lastRow As Long
    Dim lastCol As Integer
    
    lastRow = findLastRow(sheet)
    lastCol = findLastCol(sheet)
    
    If lastRow <> 0 Then
        Set selectEverythingOnSheet = sheet.Range("A1", sheet.Cells(lastRow, lastCol))
    End If

End Function

Sub copyRow(ByRef sourceSheet As Excel.Worksheet, ByVal sourceRow As Integer, ByRef destinationSheet As Excel.Worksheet, ByVal destinationRow As Integer)
    sourceSheet.Rows(sourceRow).Copy Destination:=destinationSheet.Rows(destinationRow)
End Sub

Function nameFromPath(ByVal path As String) As String

Dim parts() As String
Dim sep As String
Dim sep2 As String

sep = "\"
sep2 = "/"
If InStr(path, sep) = 0 And InStr(path, sep2) > 0 Then
    sep = sep2
End If

parts = Split(path, sep)

If UBound(parts) > -1 Then
    nameFromPath = parts(UBound(parts))
End If

End Function
