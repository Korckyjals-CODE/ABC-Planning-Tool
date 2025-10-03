Attribute VB_Name = "Reports"
Option Explicit

' ===========================
' GenerateAttendanceReport
' ===========================
' Generates an attendance report from weekly gradebooks
' 
' Parameters:
' - strBimester: The bimester folder name (e.g., "B1", "B2", "B3", "B4")
'
' Process:
' 1) Create or clear "Attendance Report" sheet
' 2) Set up tblAttendanceReport listobject with required columns
' 3) Scan Grades folder structure for weekly gradebooks
' 4) Extract attendance data for students with "AI" or "AJ" codes
' 5) Populate the attendance table
'
' Logging:
' - Immediate window (Debug.Print)
' - Sheet "ReportsLog" in ThisWorkbook (cleared/created each run)
'
' Notes:
' - Weekly gradebooks opened with macros disabled for faster processing
' - Expected ~21 gradebooks per week × ~5 weeks = ~105 files per bimester
' - File pattern: "Weekly Grade - WXX - GRADE - YEAR.xlsx"
' - Data extracted from class worksheets ("Clase 1", "Clase 2", etc.)
'
Public Sub GenerateAttendanceReport(ByVal strBimester As String)
    
    ' Variables
    Dim logLines As Collection
    Dim fso As Object
    Dim gradesFolderPath As String
    Dim bimesterFolderPath As String
    Dim wsAttendance As Worksheet
    Dim tblAttendance As ListObject
    Dim MyProgressbar As ProgressBar
    Dim totalFiles As Long
    Dim processedFiles As Long
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim ext As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim attendanceData As Collection
    Dim rowData As Collection
    
    ' Initialize
    Set logLines = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set attendanceData = New Collection
    
    ' Store original Excel settings
    Dim originalDisplayAlerts As Boolean
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    originalDisplayAlerts = Application.DisplayAlerts
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    
    ' Disable Excel features for faster processing and to prevent dialogs
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Log logLines, "=== Starting Attendance Report Generation ==="
    Log logLines, "Bimester: " & strBimester
    Log logLines, "Timestamp: " & Now()
    
    ' Validate input
    If Len(strBimester) = 0 Then
        Log logLines, "ERROR: Bimester parameter is required"
        GoTo Cleanup
    End If
    
    ' Set up paths - use the same pattern as other modules
    ' NOTE: Update this path to match your actual Grades folder location
    gradesFolderPath = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Grades\")
    bimesterFolderPath = JoinPath(gradesFolderPath, strBimester)
    
    Log logLines, "Grades folder: " & gradesFolderPath
    Log logLines, "Bimester folder: " & bimesterFolderPath
    
    ' Validate folder exists
    If Not fso.FolderExists(gradesFolderPath) Then
        Log logLines, "ERROR: Grades folder not found: " & gradesFolderPath
        GoTo Cleanup
    End If
    
    If Not fso.FolderExists(bimesterFolderPath) Then
        Log logLines, "ERROR: Bimester folder not found: " & bimesterFolderPath
        GoTo Cleanup
    End If
    
    ' Create or clear Attendance Report sheet
    Set wsAttendance = GetOrCreateWorksheet("Attendance Report")
    wsAttendance.Cells.Clear
    Log logLines, "Created/cleared Attendance Report sheet"
    
    ' Create tblAttendanceReport listobject
    Set tblAttendance = CreateAttendanceTable(wsAttendance)
    Log logLines, "Created tblAttendanceReport listobject"
    
    ' Count total files for progress tracking
    Set folder = fso.GetFolder(bimesterFolderPath)
    totalFiles = 0
    
    For Each subfolder In folder.SubFolders
        For Each file In subfolder.Files
            ext = LCase(fso.GetExtensionName(file.Name))
            If ext = "xlsx" Or ext = "xlsm" Then
                If InStr(1, file.Name, "Weekly Grade - W", vbTextCompare) > 0 Then
                    totalFiles = totalFiles + 1
                End If
            End If
        Next file
    Next subfolder
    
    Log logLines, "Found " & totalFiles & " weekly gradebook files to process"
    
    ' Initialize ProgressBar
    If totalFiles > 0 Then
        Set MyProgressbar = New ProgressBar
        
        With MyProgressbar
            .Title = "Generating Attendance Report - " & strBimester
            .ExcelStatusBar = True
            .StartColour = rgbMediumSeaGreen
            .EndColour = rgbGreen
            .TotalActions = totalFiles
        End With
        
        MyProgressbar.ShowBar
    End If
    
    ' Process weekly gradebooks
    processedFiles = 0
    For Each subfolder In folder.SubFolders
        Log logLines, "Processing subfolder: " & subfolder.Name
        For Each file In subfolder.Files
            ext = LCase(fso.GetExtensionName(file.Name))
            If ext = "xlsx" Or ext = "xlsm" Then
                If InStr(1, file.Name, "Weekly Grade - W", vbTextCompare) > 0 Then
                    processedFiles = processedFiles + 1
                    MyProgressbar.NextAction "Processing '" & file.Name & "'", True
                    Log logLines, "Found weekly gradebook: " & file.Name
                    
                    ' Extract attendance data from this file
                    ExtractAttendanceFromFile file.path, attendanceData, logLines
                End If
            End If
        Next file
    Next subfolder
    
    Log logLines, "Processed " & processedFiles & " weekly gradebook files"
    
    ' Populate attendance table
    Log logLines, "Populating attendance table with " & attendanceData.Count & " records"
    If attendanceData.Count > 0 Then
        PopulateAttendanceTable tblAttendance, attendanceData, logLines
    Else
        Log logLines, "WARNING: No attendance data found. Check the log above for issues with file processing."
    End If
    
    ' Autofit columns
    AutoFitColumnsInListObjects
    
    ' Complete progress bar
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Complete 2, "Attendance report generated successfully"
    End If
    
    Log logLines, "=== Attendance Report Generation Complete ==="
    Log logLines, "Total records: " & attendanceData.Count
    
Cleanup:
    ' Restore Excel settings
    Application.DisplayAlerts = originalDisplayAlerts
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    
    ' Log to ReportsLog sheet
    LogToReportsSheet logLines
    
    ' Clean up objects
    Set logLines = Nothing
    Set fso = Nothing
    Set wsAttendance = Nothing
    Set tblAttendance = Nothing
    Set MyProgressbar = Nothing
    Set attendanceData = Nothing
    
End Sub

' ===========================
' Helper Functions
' ===========================

Private Function GetOrCreateWorksheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If
    
    Set GetOrCreateWorksheet = ws
End Function

Private Function CreateAttendanceTable(ByVal ws As Worksheet) As ListObject
    Dim tbl As ListObject
    Dim rng As Range
    
    ' Set up headers
    ws.Cells(1, 1).Value = "Nombre"
    ws.Cells(1, 2).Value = "Grado"
    ws.Cells(1, 3).Value = "Semana de inasistencia"
    ws.Cells(1, 4).Value = "Clase de inasistencia"
    ws.Cells(1, 5).Value = "Tipo de inasistencia"
    ws.Cells(1, 6).Value = "Actividad de clase"
    
    ' Create listobject
    Set rng = ws.Range("A1:F1")
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.Name = "tblAttendanceReport"
    
    Set CreateAttendanceTable = tbl
End Function

Private Sub ExtractAttendanceFromFile(ByVal filePath As String, ByRef attendanceData As Collection, ByRef logLines As Collection)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fso As Object
    Dim fileName As String
    Dim grade As String
    Dim week As String
    Dim activity As String
    Dim i As Long
    Dim rowData As Collection
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(filePath)
    
    Log logLines, "Processing file: " & fileName
    
    ' Extract grade and week from filename
    grade = ExtractGradeFromFileName(fileName)
    week = ExtractWeekFromFileName(fileName)
    
    If Len(grade) = 0 Or Len(week) = 0 Then
        Log logLines, "WARNING: Could not extract grade/week from: " & fileName
        Exit Sub
    End If
    
    ' Store original Excel settings
    Dim originalDisplayAlerts As Boolean
    Dim originalScreenUpdating As Boolean
    originalDisplayAlerts = Application.DisplayAlerts
    originalScreenUpdating = Application.ScreenUpdating
    
    ' Disable Excel features to prevent message boxes
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ' Open workbook with macros disabled
    On Error Resume Next
    Set wb = Application.Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True)
    If Err.Number <> 0 Then
        Log logLines, "ERROR: Could not open file: " & fileName & " | " & Err.Description
        Err.Clear
        ' Restore Excel settings before exit
        Application.DisplayAlerts = originalDisplayAlerts
        Application.ScreenUpdating = originalScreenUpdating
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Restore Excel settings
    Application.DisplayAlerts = originalDisplayAlerts
    Application.ScreenUpdating = originalScreenUpdating
    
    ' Process each worksheet
    For Each ws In wb.Worksheets
        If InStr(1, ws.Name, "Clase ", vbTextCompare) > 0 Then
            ' Get activity by finding "Contexto" or "Objetivo" and taking the cell to its right
            activity = GetClassActivity(ws)
            If Len(activity) = 0 Then activity = "No activity specified"
            
            ' Check if Nombre and Observaciones headers exist
            Dim nombreHeader As Range
            Dim obsHeader As Range
            
            Set nombreHeader = ws.Cells.Find(What:="Nombre", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            Set obsHeader = ws.Cells.Find(What:="Observaciones", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
            
            If nombreHeader Is Nothing Then
                Log logLines, "WARNING: 'Nombre' header not found in '" & ws.Name & "'"
            ElseIf obsHeader Is Nothing Then
                Log logLines, "WARNING: 'Observaciones' header not found in '" & ws.Name & "'"
            Else
                ' Get column indices
                Dim asistenciaCol As Long
                Dim nombreCol As Long
                asistenciaCol = GetColumnIndex(ws, "Asistencia")
                nombreCol = GetColumnIndex(ws, "Nombre")
                
                Dim dataRowCount As Long
                dataRowCount = GetDataRowCount(ws)
                
                Log logLines, "Processing '" & ws.Name & "' - Data Rows: " & dataRowCount & " - Asistencia Col: " & asistenciaCol & " - Nombre Col: " & nombreCol
                
                ' Log available columns for debugging
                Dim headers As Collection
                Set headers = GetColumnHeaders(ws)
                Log logLines, "Available columns in '" & ws.Name & "':"
                For i = 1 To headers.Count
                    Log logLines, "  Column " & i & ": '" & headers(i) & "'"
                Next i
                
                If asistenciaCol = 0 Then
                    Log logLines, "WARNING: 'Asistencia' column not found in '" & ws.Name & "'"
                ElseIf nombreCol = 0 Then
                    Log logLines, "WARNING: 'Nombre' column not found in '" & ws.Name & "'"
                ElseIf dataRowCount = 0 Then
                    Log logLines, "WARNING: No data rows found in '" & ws.Name & "'"
                Else
                    ' Process each row in the data table
                    For i = 1 To dataRowCount
                        Dim asistenciaValue As String
                        Dim studentName As String
                        
                        asistenciaValue = GetCellValue(ws, i, asistenciaCol)
                        studentName = GetCellValue(ws, i, nombreCol)
                        
                        ' Log first few rows for debugging
                        If i <= 3 Then
                            Log logLines, "  Row " & i & ": Nombre='" & studentName & "' Asistencia='" & asistenciaValue & "'"
                        End If
                        
                        ' Check if student has AI or AJ
                        If asistenciaValue = "AI" Or asistenciaValue = "AJ" Then
                            Log logLines, "  FOUND ABSENCE: Row " & i & " - " & studentName & " (" & asistenciaValue & ")"
                            
                            If Len(studentName) > 0 Then
                                ' Create row data
                                Set rowData = New Collection
                                rowData.Add studentName                    ' Nombre
                                rowData.Add grade                          ' Grado
                                rowData.Add week                           ' Semana de inasistencia
                                rowData.Add ws.Name                        ' Clase de inasistencia
                                rowData.Add asistenciaValue                ' Tipo de inasistencia
                                rowData.Add activity                       ' Actividad de clase
                                
                                attendanceData.Add rowData
                            End If
                        End If
                    Next i
                End If
            End If
        End If
    Next ws
    
    ' Close workbook
    wb.Close False
    
    Log logLines, "Extracted data from: " & fileName
End Sub

Private Function ExtractGradeFromFileName(ByVal fileName As String) As String
    ' Extract grade from "Weekly Grade - WXX - GRADE - YEAR.xlsx"
    ' Pattern: between " - " and " - "
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, fileName, " - ", vbTextCompare)
    If startPos > 0 Then
        startPos = startPos + 3 ' Move past " - "
        endPos = InStr(startPos, fileName, " - ", vbTextCompare)
        If endPos > 0 Then
            ExtractGradeFromFileName = Mid(fileName, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function ExtractWeekFromFileName(ByVal fileName As String) As String
    ' Extract week from "Weekly Grade - WXX - GRADE - YEAR.xlsx"
    ' Pattern: between "Grade - " and " - "
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, fileName, "Grade - ", vbTextCompare)
    If startPos > 0 Then
        startPos = startPos + 8 ' Move past "Grade - "
        endPos = InStr(startPos, fileName, " - ", vbTextCompare)
        If endPos > 0 Then
            ExtractWeekFromFileName = Mid(fileName, startPos, endPos - startPos)
        End If
    End If
End Function

Private Function GetClassActivity(ByVal ws As Worksheet) As String
    Dim rng As Range
    Dim cell As Range
    Dim searchTerms As Variant
    Dim term As Variant
    Dim foundCell As Range
    
    ' Define search terms
    searchTerms = Array("Contexto", "Objetivo")
    
    ' Search in the first few rows and columns for efficiency
    Set rng = ws.Range("A1:Z10")
    
    For Each term In searchTerms
        Set foundCell = rng.Find(What:=term, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not foundCell Is Nothing Then
            ' Get the cell to the right of the found cell
            GetClassActivity = Trim(foundCell.Offset(0, 1).Value)
            Exit Function
        End If
    Next term
    
    ' If not found, return empty string
    GetClassActivity = ""
End Function

' ===========================
' Helper Functions for Regular Cell Range Tables
' ===========================

Private Function GetTableRange(ByVal ws As Worksheet) As Range
    ' Returns the entire data table range (headers + data rows)
    Dim nombreStart As Range
    Dim obsEnd As Range
    Dim lastDataRow As Long
    Dim tableRange As Range
    
    ' Find "Nombre" header
    Set nombreStart = ws.Cells.Find(What:="Nombre", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If nombreStart Is Nothing Then
        Set GetTableRange = Nothing
        Exit Function
    End If
    
    ' Find "Observaciones" header
    Set obsEnd = ws.Cells.Find(What:="Observaciones", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If obsEnd Is Nothing Then
        Set GetTableRange = Nothing
        Exit Function
    End If
    
    ' Find last row with data in Nombre column
    lastDataRow = GetLastDataRow(ws, nombreStart.Column)
    
    ' Create table range from Nombre column to Observaciones column, from header row to last data row
    Set tableRange = ws.Range(nombreStart.Offset(1, 0), ws.Cells(lastDataRow, obsEnd.Column))
    Set GetTableRange = tableRange
End Function

Private Function GetHeaderRow(ByVal ws As Worksheet) As Range
    ' Returns the header row (row containing "Nombre" through "Observaciones")
    Dim nombreStart As Range
    Dim obsEnd As Range
    Dim headerRow As Range
    
    Set nombreStart = ws.Cells.Find(What:="Nombre", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If nombreStart Is Nothing Then
        Set GetHeaderRow = Nothing
        Exit Function
    End If
    
    Set obsEnd = ws.Cells.Find(What:="Observaciones", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If obsEnd Is Nothing Then
        Set GetHeaderRow = Nothing
        Exit Function
    End If
    
    ' Return header row from Nombre to Observaciones
    Set headerRow = ws.Range(nombreStart, obsEnd)
    Set GetHeaderRow = headerRow
End Function

Private Function GetColumnIndex(ByVal ws As Worksheet, ByVal columnName As String) As Long
    ' Alternative version that works with regular ranges instead of ListObjects
    Dim headerRow As Range
    Dim i As Long
    
    Set headerRow = GetHeaderRow(ws)
    If headerRow Is Nothing Then
        GetColumnIndex = 0
        Exit Function
    End If
    
    ' Search through header cells
    For i = 1 To headerRow.Columns.Count
        If Trim(headerRow.Cells(1, i).Value) = columnName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    
    GetColumnIndex = 0 ' Column not found
End Function

Private Function GetDataRowCount(ByVal ws As Worksheet) As Long
    ' Returns the number of data rows (excluding header)
    Dim tableRange As Range
    
    Set tableRange = GetTableRange(ws)
    If tableRange Is Nothing Then
        GetDataRowCount = 0
        Exit Function
    End If
    
    GetDataRowCount = tableRange.Rows.Count
End Function

Private Function GetCellValue(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String
    ' Gets cell value from the data table, adjusted for relative positioning
    ' rowIndex and colIndex are relative to the data table (1-based, excluding headers)
    Dim nombreStart As Range
    
    Set nombreStart = ws.Cells.Find(What:="Nombre", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If nombreStart Is Nothing Then
        GetCellValue = ""
        Exit Function
    End If
    
    ' Calculate absolute cell position
    ' rowIndex = 1 means first data row (below header), so we add 1 to skip header
    ' colIndex = 1 means Nombre column (leftmost), so we add column offset
    Dim absRow As Long
    Dim absCol As Long
    
    absRow = nombreStart.Row + rowIndex ' Skip header row
    absCol = nombreStart.Column + colIndex - 1 ' Adjust for relative column position
    
    GetCellValue = Trim(GetActualValue(ws.Cells(absRow, absCol)))
End Function

Private Function GetColumnHeaders(ByVal ws As Worksheet) As Collection
    ' Returns a collection of column header names
    Dim headers As Collection
    Dim headerRow As Range
    Dim i As Long
    
    Set headers = New Collection
    Set headerRow = GetHeaderRow(ws)
    
    If headerRow Is Nothing Then
        Set GetColumnHeaders = headers
        Exit Function
    End If
    
    ' Add each header to collection
    For i = 1 To headerRow.Columns.Count
        headers.Add Trim(headerRow.Cells(1, i).Value)
    Next i
    
    Set GetColumnHeaders = headers
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet, ByVal column As Long) As Long
    ' Finds the last row with non-empty data in the specified column
    ' Excludes rows where Value is 0 (from empty cells)
    Dim lastRow As Long
    Dim currentRow As Long
    Dim currentValue As Variant
    
    lastRow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
    
    ' Work backwards to find the actual last data row
    For currentRow = lastRow To 2 Step -1 ' Start from bottom, go up to row 2 (skip header)
        currentValue = GetActualValue(ws.Cells(currentRow, column))
        
        ' Check if value is meaningful (not empty, not 0)
        If IsNotEmptyOrZero(currentValue) Then
            GetLastDataRow = currentRow
            Exit Function
        End If
    Next currentRow
    
    ' If no data found, return header row
    GetLastDataRow = 1
End Function

Private Function GetActualValue(ByVal cell As Range) As Variant
    ' Gets the actual value from a cell, handling cases where Value returns 0 for empty cells
    If cell.Text = "" Then
        GetActualValue = ""
    ElseIf IsNumeric(cell.Value) And cell.Value = 0 And cell.Text = "0" Then
        ' This is likely a real zero value
        GetActualValue = cell.Value
    ElseIf IsNumeric(cell.Value) And cell.Value = 0 And cell.Text = "" Then
        ' This is an empty cell that Excel returns as 0
        GetActualValue = ""
    Else
        GetActualValue = cell.Value
    End If
End Function

Private Function IsNotEmptyOrZero(value As Variant) As Boolean
    ' Checks if a value is not empty and not the 0 that comes from empty cells
    If IsEmpty(value) Then
        IsNotEmptyOrZero = False
    ElseIf Len(value & "") = 0 Then
        IsNotEmptyOrZero = False
    ElseIf IsNumeric(value) And value = 0 Then
        IsNotEmptyOrZero = False
    Else
        IsNotEmptyOrZero = True
    End If
End Function

Private Sub PopulateAttendanceTable(ByVal tbl As ListObject, ByVal attendanceData As Collection, ByRef logLines As Collection)
    Dim i As Long
    Dim rowData As Collection
    Dim newRow As ListRow
    
    For i = 1 To attendanceData.Count
        Set rowData = attendanceData(i)
        
        ' Add new row to table
        Set newRow = tbl.ListRows.Add
        
        ' Populate row data
        newRow.Range.Cells(1, 1).Value = rowData(1) ' Nombre
        newRow.Range.Cells(1, 2).Value = rowData(2) ' Grado
        newRow.Range.Cells(1, 3).Value = rowData(3) ' Semana de inasistencia
        newRow.Range.Cells(1, 4).Value = rowData(4) ' Clase de inasistencia
        newRow.Range.Cells(1, 5).Value = rowData(5) ' Tipo de inasistencia
        newRow.Range.Cells(1, 6).Value = rowData(6) ' Actividad de clase
    Next i
    
    Log logLines, "Added " & attendanceData.Count & " rows to attendance table"
End Sub

Private Sub LogToReportsSheet(ByVal logLines As Collection)
    Dim ws As Worksheet
    Dim i As Long
    
    ' Create or clear ReportsLog sheet
    Set ws = GetOrCreateWorksheet("ReportsLog")
    ws.Cells.Clear
    
    ' Add headers
    ws.Cells(1, 1).Value = "Timestamp"
    ws.Cells(1, 2).Value = "Log Entry"
    ws.Range("A1:B1").Font.Bold = True
    
    ' Check if logLines collection is valid and has items
    If logLines Is Nothing Then
        ws.Cells(2, 1).Value = EscapeFormula("Error")
        ws.Cells(2, 2).Value = EscapeFormula("LogLines collection is Nothing")
        Exit Sub
    End If
    
    If logLines.Count = 0 Then
        ws.Cells(2, 1).Value = EscapeFormula("Info")
        ws.Cells(2, 2).Value = EscapeFormula("No log entries found")
        Exit Sub
    End If
    
    ' Add log entries with error handling
    For i = 1 To logLines.Count
        On Error Resume Next
        Dim logEntry As String
        Dim timestamp As String
        Dim message As String
        
        ' Safely get log entry from collection
        logEntry = logLines(i)
        If Err.Number <> 0 Then
            ws.Cells(i + 1, 1).Value = EscapeFormula("Error")
            ws.Cells(i + 1, 2).Value = EscapeFormula("Could not access log entry " & i & ": " & Err.Description)
            Err.Clear
            GoTo NextLogEntry
        End If
        
        ' Check if logEntry is valid
        If Len(Trim(logEntry)) = 0 Then
            logEntry = ""
        End If
        
        ' Check if the log entry contains the expected format "timestamp  message"
        If InStr(1, logEntry, "  ") > 0 Then
            ' Split timestamp and message
            Dim parts() As String
            parts = Split(logEntry, "  ", 2) ' Split into max 2 parts
            
            If UBound(parts) >= 0 Then
                timestamp = parts(0)
            Else
                timestamp = "Unknown"
            End If
            
            If UBound(parts) >= 1 Then
                message = parts(1)
            Else
                message = logEntry
            End If
        Else
            ' If no separator found, treat entire entry as message
            timestamp = "Unknown"
            message = logEntry
        End If
        
        ' Safely write to worksheet - escape formulas by adding single quote
        ws.Cells(i + 1, 1).Value = EscapeFormula(timestamp)
        ws.Cells(i + 1, 2).Value = EscapeFormula(message)
        
NextLogEntry:
        On Error GoTo 0
    Next i
    
    ' Autofit columns
    ws.Columns("A:B").EntireColumn.AutoFit
End Sub

Private Function EscapeFormula(ByVal value As String) As String
    ' Escape values that start with = to prevent Excel from interpreting them as formulas
    If Left$(value, 1) = "=" Then
        EscapeFormula = "'" & value
    Else
        EscapeFormula = value
    End If
End Function

Private Function TrimTrailingSlash(ByVal p As String) As String
    If Right$(p, 1) = "\" Then
        TrimTrailingSlash = Left$(p, Len(p) - 1)
    Else
        TrimTrailingSlash = p
    End If
End Function

Private Function JoinPath(ByVal base As String, ByVal leaf As String) As String
    JoinPath = TrimTrailingSlash(base) & "\" & leaf
End Function

Private Sub Log(ByRef logLines As Collection, ByVal msg As String)
    Dim timestampedMsg As String
    timestampedMsg = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    
    ' Add to collection for worksheet logging
    logLines.Add timestampedMsg
    
    ' Also print to immediate window
    Debug.Print timestampedMsg
End Sub
