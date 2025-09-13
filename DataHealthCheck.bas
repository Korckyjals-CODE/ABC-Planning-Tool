Attribute VB_Name = "DataHealthCheck"
Option Explicit

' ===========================
' Data Health Check Module
' Milestone 1: Basic Health Check Framework + Enhanced Reporting
' ===========================

' Constants
Private Const DEFAULT_GRADE_VALUE As Integer = 20
Private Const EXPECTED_WEIGHT_SUM As Integer = 100
Private Const HEALTH_REPORT_SHEET_NAME As String = "HealthReport"

' ===========================
' Main Health Check Functions
' ===========================

Public Sub RunBasicHealthCheck()
    ' Main entry point for health check
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If IsWeeklyGradebook(ws) Then
        ValidateWeeklyGradebookBasic ws
    Else
        MsgBox "Please select a weekly gradebook sheet to run health check.", vbExclamation, "Health Check"
    End If
End Sub

Public Sub RunHealthCheckOnWorkbook(ByVal wb As Workbook)
    ' Run health check on a specific workbook (for external gradebooks)
    Dim ws As Worksheet
    Dim issuesFound As Boolean
    issuesFound = False
    
    For Each ws In wb.Worksheets
        If IsWeeklyGradebook(ws) Then
            Debug.Print "Running health check on: " & wb.Name & " - " & ws.Name
            ValidateWeeklyGradebookBasic ws
            issuesFound = True
        End If
    Next ws
    
    If Not issuesFound Then
        MsgBox "No weekly gradebook sheets found in: " & wb.Name, vbInformation, "Health Check"
    End If
End Sub

Public Sub RunHealthCheckOnFile(ByVal filePath As String)
    ' Run health check on a specific file path
    Dim wb As Workbook
    Dim xlApp As Object
    Dim xlAppCreated As Boolean
    
    ' Try to get existing Excel instance first
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Set xlApp = CreateObject("Excel.Application")
        xlAppCreated = True
        xlApp.Visible = False
    End If
    On Error GoTo 0
    
    If xlApp Is Nothing Then
        MsgBox "Could not access Excel application to open file: " & filePath, vbExclamation, "Health Check"
        Exit Sub
    End If
    
    ' Open the workbook
    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(filePath)
    If Err.Number <> 0 Then
        MsgBox "Could not open file: " & filePath & vbCrLf & "Error: " & Err.Description, vbExclamation, "Health Check"
        Err.Clear
        If xlAppCreated Then xlApp.Quit
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Run health check
    RunHealthCheckOnWorkbook wb
    
    ' Close workbook and Excel instance if we created it
    wb.Close SaveChanges:=False
    If xlAppCreated Then xlApp.Quit
End Sub

Public Sub RunHealthCheckOnSheet(ByVal ws As Worksheet)
    ' Run health check on specific worksheet
    If IsWeeklyGradebook(ws) Then
        ValidateWeeklyGradebookBasic ws
    Else
        MsgBox "Selected sheet is not a recognized weekly gradebook format.", vbExclamation, "Health Check"
    End If
End Sub

' ===========================
' Weekly Gradebook Detection
' ===========================

Private Function IsWeeklyGradebook(ByVal ws As Worksheet) As Boolean
    ' Detect if worksheet is a weekly gradebook based on structure
    Dim hasWeeklyStructure As Boolean
    hasWeeklyStructure = False
    
    ' Check for common weekly gradebook indicators
    If ws.Cells(1, 1).Value Like "*Nota Semanal*" Or _
       ws.Cells(1, 3).Value Like "*Nota Semanal*" Then
        hasWeeklyStructure = True
    End If
    
    ' Check for class columns (Clase 1, Clase 2, etc.)
    If Not hasWeeklyStructure Then
        Dim col As Integer
        For col = 3 To 7 ' Check columns C through G
            If ws.Cells(3, col).Value Like "Clase *" Then
                hasWeeklyStructure = True
                Exit For
            End If
        Next col
    End If
    
    ' Check for percentage weights in row 2
    If Not hasWeeklyStructure Then
        Dim hasWeights As Boolean
        hasWeights = False
        For col = 3 To 7
            If IsNumeric(ws.Cells(2, col).Value) And ws.Cells(2, col).Value Like "*%" Then
                hasWeights = True
                Exit For
            End If
        Next col
        hasWeeklyStructure = hasWeights
    End If
    
    IsWeeklyGradebook = hasWeeklyStructure
End Function

' ===========================
' Basic Weekly Gradebook Validation
' ===========================

Private Sub ValidateWeeklyGradebookBasic(ByVal ws As Worksheet)
    Dim issues As Collection
    Set issues = New Collection
    
    ' Log start of validation
    Debug.Print "Starting health check for: " & ws.Name
    
    ' Check for classes with default grades but non-zero weight
    CheckDefaultGradesWithWeight ws, issues
    
    ' Check weight sum
    CheckWeightSum ws, issues
    
    ' Check for empty student rows with default grades
    CheckEmptyStudentRows ws, issues
    
    ' Report results to health report sheet
    ReportHealthIssuesToSheet issues, ws.Name, ws.Parent.Name
End Sub

Private Sub CheckDefaultGradesWithWeight(ByVal ws As Worksheet, ByRef issues As Collection)
    ' Check if any class has all default grades (20) but non-zero weight
    Dim classColumns As Collection
    Set classColumns = GetClassColumns(ws)
    
    Dim classCol As Variant
    For Each classCol In classColumns
        Dim colNum As Integer
        colNum = CInt(classCol)
        
        If AllGradesEqualDefault(ws, colNum) And GetClassWeight(ws, colNum) > 0 Then
            Dim issue As String
            issue = "Class " & GetClassNumber(ws, colNum) & " has all default grades (20) but weight > 0%. " & _
                   "Consider setting weight to 0% if there was no class."
            issues.Add issue
            Debug.Print "ISSUE: " & issue
        End If
    Next classCol
End Sub

Private Sub CheckWeightSum(ByVal ws As Worksheet, ByRef issues As Collection)
    ' Check if weights sum to 100%
    Dim totalWeight As Double
    totalWeight = 0
    
    Dim classColumns As Collection
    Set classColumns = GetClassColumns(ws)
    
    Dim classCol As Variant
    For Each classCol In classColumns
        totalWeight = totalWeight + GetClassWeight(ws, CInt(classCol))
    Next classCol
    
    If totalWeight <> EXPECTED_WEIGHT_SUM Then
        Dim issue As String
        issue = "Class weights sum to " & totalWeight & "% instead of " & EXPECTED_WEIGHT_SUM & "%"
        issues.Add issue
        Debug.Print "ISSUE: " & issue
    End If
End Sub

Private Sub CheckEmptyStudentRows(ByVal ws As Worksheet, ByRef issues As Collection)
    ' Check for student rows that have default grades but no name
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    Dim row As Long
    For row = 4 To lastRow ' Assuming student data starts at row 4
        If IsEmptyStudentRow(ws, row) Then
            Dim issue As String
            issue = "Row " & row & " appears to be empty student data with default grades"
            issues.Add issue
            Debug.Print "ISSUE: " & issue
        End If
    Next row
End Sub

' ===========================
' Helper Functions
' ===========================

Private Function GetClassColumns(ByVal ws As Worksheet) As Collection
    ' Get collection of class column numbers (C, D, E, F, G)
    Dim classCols As Collection
    Set classCols = New Collection
    
    Dim col As Integer
    For col = 3 To 7 ' Columns C through G
        If ws.Cells(3, col).Value Like "Clase *" Then
            classCols.Add col
        End If
    Next col
    
    Set GetClassColumns = classCols
End Function

Private Function AllGradesEqualDefault(ByVal ws As Worksheet, ByVal colNum As Integer) As Boolean
    ' Check if all grades in a column are the default value (20)
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    Dim row As Long
    For row = 4 To lastRow ' Assuming student data starts at row 4
        If Not IsEmpty(ws.Cells(row, colNum).Value) Then
            If ws.Cells(row, colNum).Value <> DEFAULT_GRADE_VALUE Then
                AllGradesEqualDefault = False
                Exit Function
            End If
        End If
    Next row
    
    AllGradesEqualDefault = True
End Function

Private Function GetClassWeight(ByVal ws As Worksheet, ByVal colNum As Integer) As Double
    ' Get the weight percentage for a class column
    Dim weightStr As String
    weightStr = CStr(ws.Cells(2, colNum).Value)
    
    If weightStr Like "*%" Then
        GetClassWeight = CDbl(Replace(weightStr, "%", ""))
    Else
        GetClassWeight = 0
    End If
End Function

Private Function GetClassNumber(ByVal ws As Worksheet, ByVal colNum As Integer) As String
    ' Extract class number from header (e.g., "Clase 1" -> "1")
    Dim header As String
    header = CStr(ws.Cells(3, colNum).Value)
    
    If header Like "Clase *" Then
        GetClassNumber = Replace(header, "Clase ", "")
    Else
        GetClassNumber = "Unknown"
    End If
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    ' Find the last row with data in column A (student names)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' If no data found, return 0
    If lastRow = 1 And IsEmpty(ws.Cells(1, 1).Value) Then
        GetLastDataRow = 0
    Else
        GetLastDataRow = lastRow
    End If
End Function

Private Function IsEmptyStudentRow(ByVal ws As Worksheet, ByVal row As Long) As Boolean
    ' Check if a row is an empty student row (no name but has default grades)
    Dim hasName As Boolean
    Dim hasDefaultGrades As Boolean
    
    ' Check if there's a name in column A
    hasName = Not IsEmpty(ws.Cells(row, 1).Value) And ws.Cells(row, 1).Value <> "0"
    
    ' Check if all class columns have default grades
    hasDefaultGrades = True
    Dim classColumns As Collection
    Set classColumns = GetClassColumns(ws)
    
    Dim classCol As Variant
    For Each classCol In classColumns
        If ws.Cells(row, CInt(classCol)).Value <> DEFAULT_GRADE_VALUE Then
            hasDefaultGrades = False
            Exit For
        End If
    Next classCol
    
    ' Empty student row = no name but has default grades
    IsEmptyStudentRow = Not hasName And hasDefaultGrades
End Function

' ===========================
' Enhanced Reporting Functions
' ===========================

Private Sub ReportHealthIssuesToSheet(ByVal issues As Collection, ByVal sheetName As String, ByVal workbookName As String)
    ' Create or update health report sheet with detailed results
    Dim reportWs As Worksheet
    Set reportWs = GetOrCreateHealthReportSheet
    
    ' Add new health check entry
    AddHealthCheckEntry reportWs, issues, sheetName, workbookName
    
    ' Show summary message
    If issues.Count = 0 Then
        MsgBox "✓ " & sheetName & " - No health issues found!" & vbCrLf & vbCrLf & _
               "The gradebook data appears to be healthy." & vbCrLf & vbCrLf & _
               "Check the 'HealthReport' sheet for detailed results.", vbInformation, "Health Check Complete"
    Else
        MsgBox "⚠ " & sheetName & " - " & issues.Count & " health issue(s) found!" & vbCrLf & vbCrLf & _
               "Please review the 'HealthReport' sheet for detailed information." & vbCrLf & vbCrLf & _
               "You can now work with your gradebook while keeping the report open.", vbExclamation, "Health Check Complete"
    End If
    
    ' Also log to immediate window
    Debug.Print "Health check complete for " & sheetName & " - " & issues.Count & " issues found"
End Sub

Private Function GetOrCreateHealthReportSheet() As Worksheet
    ' Get existing health report sheet or create a new one in the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = ThisWorkbook
    
    ' Try to get existing sheet
    On Error Resume Next
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    ' If sheet doesn't exist, create it
    If reportWs Is Nothing Then
        Set reportWs = planningWb.Worksheets.Add
        reportWs.Name = HEALTH_REPORT_SHEET_NAME
        SetupHealthReportSheet reportWs
    End If
    
    Set GetOrCreateHealthReportSheet = reportWs
End Function

Private Sub SetupHealthReportSheet(ByVal ws As Worksheet)
    ' Set up the health report sheet with headers and formatting
    With ws
        ' Clear any existing content
        .Cells.Clear
        
        ' Headers
        .Cells(1, 1).Value = "Data Health Check Report"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Interior.Color = RGB(68, 114, 196)
        .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        
        ' Subtitle
        .Cells(2, 1).Value = "Generated on: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(2, 1).Font.Italic = True
        
        ' Column headers
        .Cells(4, 1).Value = "Timestamp"
        .Cells(4, 2).Value = "Workbook"
        .Cells(4, 3).Value = "Sheet"
        .Cells(4, 4).Value = "Status"
        .Cells(4, 5).Value = "Issues Count"
        .Cells(4, 6).Value = "Issue Details"
        
        ' Format headers
        Dim headerRange As Range
        Set headerRange = .Range("A4:F4")
        With headerRange
            .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        
        ' Set column widths
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 25
        .Columns("C").ColumnWidth = 20
        .Columns("D").ColumnWidth = 15
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 60
        
        ' Freeze panes
        .Range("A5").Select
        ActiveWindow.FreezePanes = True
        
        ' Add instructions
        .Cells(6, 1).Value = "Instructions:"
        .Cells(6, 1).Font.Bold = True
        .Cells(7, 1).Value = "• Green status = No issues found"
        .Cells(8, 1).Value = "• Red status = Issues found - review details in column F"
        .Cells(9, 1).Value = "• Click on any row to see more details"
        .Cells(10, 1).Value = "• This report updates automatically with each health check"
        
        ' Format instruction text
        .Range("A7:A10").Font.Size = 10
        .Range("A7:A10").Font.Color = RGB(100, 100, 100)
    End With
End Sub

Private Sub AddHealthCheckEntry(ByVal reportWs As Worksheet, ByVal issues As Collection, ByVal sheetName As String, ByVal workbookName As String)
    ' Add a new health check entry to the report sheet
    Dim lastRow As Long
    lastRow = reportWs.Cells(reportWs.Rows.Count, 1).End(xlUp).Row
    
    ' Find next available row (skip instruction rows)
    If lastRow < 10 Then lastRow = 10
    
    Dim newRow As Long
    newRow = lastRow + 1
    
    ' Add basic information
    reportWs.Cells(newRow, 1).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
    reportWs.Cells(newRow, 2).Value = workbookName
    reportWs.Cells(newRow, 3).Value = sheetName
    reportWs.Cells(newRow, 5).Value = issues.Count
    
    ' Add status and formatting
    If issues.Count = 0 Then
        reportWs.Cells(newRow, 4).Value = "✓ HEALTHY"
        reportWs.Cells(newRow, 4).Interior.Color = RGB(198, 239, 206)
        reportWs.Cells(newRow, 4).Font.Color = RGB(0, 100, 0)
        reportWs.Cells(newRow, 6).Value = "No issues detected"
    Else
        reportWs.Cells(newRow, 4).Value = "⚠ ISSUES"
        reportWs.Cells(newRow, 4).Interior.Color = RGB(255, 199, 206)
        reportWs.Cells(newRow, 4).Font.Color = RGB(156, 0, 6)
        
        ' Add detailed issue descriptions
        Dim issueDetails As String
        issueDetails = ""
        Dim i As Integer
        For i = 1 To issues.Count
            issueDetails = issueDetails & i & ". " & issues(i)
            If i < issues.Count Then issueDetails = issueDetails & vbLf
        Next i
        reportWs.Cells(newRow, 6).Value = issueDetails
        reportWs.Cells(newRow, 6).WrapText = True
    End If
    
    ' Add borders
    Dim dataRange As Range
    Set dataRange = reportWs.Range("A" & newRow & ":F" & newRow)
    With dataRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Auto-fit row height for issue details
    reportWs.Rows(newRow).AutoFit
    
    ' Scroll to the new entry
    reportWs.Cells(newRow, 1).Select
End Sub

' ===========================
' Legacy Reporting Functions (for backward compatibility)
' ===========================

Private Sub ReportHealthIssues(ByVal issues As Collection, ByVal sheetName As String)
    ' Legacy function - now redirects to sheet-based reporting
    ReportHealthIssuesToSheet issues, sheetName, ThisWorkbook.Name
End Sub

' ===========================
' Utility Functions
' ===========================

Public Sub TestHealthCheck()
    ' Test function to validate health check on current sheet
    Debug.Print "Testing health check on: " & ActiveSheet.Name
    RunBasicHealthCheck
End Sub

Public Sub ShowGradebookInfo()
    ' Show information about the current gradebook structure
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Debug.Print "=== Gradebook Information ==="
    Debug.Print "Sheet Name: " & ws.Name
    Debug.Print "Is Weekly Gradebook: " & IsWeeklyGradebook(ws)
    
    If IsWeeklyGradebook(ws) Then
        Dim classColumns As Collection
        Set classColumns = GetClassColumns(ws)
        
        Debug.Print "Class Columns Found: " & classColumns.Count
        Dim classCol As Variant
        For Each classCol In classColumns
            Debug.Print "  Column " & classCol & ": " & ws.Cells(3, CInt(classCol)).Value & " (Weight: " & GetClassWeight(ws, CInt(classCol)) & "%)"
        Next classCol
        
        Debug.Print "Last Data Row: " & GetLastDataRow(ws)
    End If
    Debug.Print "============================="
End Sub

' ===========================
' Integration with GenerateRawGradebooks
' ===========================

Public Sub RunHealthCheckOnGeneratedGradebook(ByVal wb As Object, ByVal templatePath As String)
    ' Run health check on a gradebook generated by GenerateRawGradebooks
    ' This function is designed to be called from UpdateGradebooks.bas
    
    If wb Is Nothing Then
        Debug.Print "Health Check: Cannot run on Nothing workbook - " & templatePath
        Exit Sub
    End If
    
    ' Check if workbook is still accessible
    On Error Resume Next
    Dim testCount As Long
    testCount = wb.Worksheets.Count
    If Err.Number <> 0 Then
        Debug.Print "Health Check: Workbook no longer accessible - " & templatePath
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Run health check on the first worksheet (gradebooks typically have one sheet)
    Dim ws As Object
    Set ws = wb.Worksheets(1)
    
    If IsWeeklyGradebook(ws) Then
        Debug.Print "Health Check: Running on generated gradebook - " & templatePath
        ValidateWeeklyGradebookBasic ws
    Else
        Debug.Print "Health Check: Generated gradebook is not recognized format - " & templatePath
    End If
End Sub

Public Sub RunHealthCheckOnFolder(ByVal folderPath As String, Optional ByVal bimester As String = "")
    ' Run health check on all gradebook files in a folder
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim processedCount As Integer
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbExclamation, "Health Check"
        Exit Sub
    End If
    
    Set folder = fso.GetFolder(folderPath)
    processedCount = 0
    
    Debug.Print "=== Starting Health Check on Folder ==="
    Debug.Print "Folder: " & folderPath
    If bimester <> "" Then Debug.Print "Bimester: " & bimester
    Debug.Print "======================================="
    
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xlsx" Then
            If bimester = "" Or InStr(1, file.Name, bimester, vbTextCompare) > 0 Then
                Debug.Print "Processing: " & file.Name
                RunHealthCheckOnFile file.Path
                processedCount = processedCount + 1
            End If
        End If
    Next file
    
    Debug.Print "=== Health Check Complete ==="
    Debug.Print "Files processed: " & processedCount
    Debug.Print "============================="
    
    MsgBox "Health check completed on " & processedCount & " files in folder: " & folderPath, vbInformation, "Health Check Complete"
End Sub

' ===========================
' Health Report Management
' ===========================

Public Sub ClearHealthReport()
    ' Clear the health report sheet from the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = ThisWorkbook
    
    On Error Resume Next
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    If Not reportWs Is Nothing Then
        If MsgBox("Are you sure you want to clear the health report?", vbYesNo + vbQuestion, "Clear Health Report") = vbYes Then
            reportWs.Delete
            MsgBox "Health report cleared from planning workbook.", vbInformation, "Health Report"
        End If
    Else
        MsgBox "No health report found to clear in planning workbook.", vbInformation, "Health Report"
    End If
End Sub

Public Sub ShowHealthReport()
    ' Show the health report sheet from the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = ThisWorkbook
    
    On Error Resume Next
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    If Not reportWs Is Nothing Then
        ' Switch to planning workbook and activate the report sheet
        planningWb.Activate
        reportWs.Activate
        reportWs.Cells(1, 1).Select
    Else
        MsgBox "No health report found in planning workbook. Run a health check first.", vbInformation, "Health Report"
    End If
End Sub