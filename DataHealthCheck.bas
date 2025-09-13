Attribute VB_Name = "DataHealthCheck"
Option Explicit

' ===========================
' Data Health Check Module
' Milestone 1: Basic Health Check Framework
' ===========================

' Constants
Private Const DEFAULT_GRADE_VALUE As Integer = 20
Private Const EXPECTED_WEIGHT_SUM As Integer = 100

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
    
    ' Report results
    ReportHealthIssues issues, ws.Name
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
' Reporting Functions
' ===========================

Private Sub ReportHealthIssues(ByVal issues As Collection, ByVal sheetName As String)
    ' Report health check results to user
    If issues.Count = 0 Then
        MsgBox "✓ " & sheetName & " - No health issues found!" & vbCrLf & vbCrLf & _
               "The gradebook data appears to be healthy.", vbInformation, "Health Check Complete"
    Else
        Dim message As String
        message = "⚠ " & sheetName & " - " & issues.Count & " health issue(s) found:" & vbCrLf & vbCrLf
        
        Dim i As Integer
        For i = 1 To issues.Count
            message = message & i & ". " & issues(i) & vbCrLf
        Next i
        
        message = message & vbCrLf & "Please review and correct these issues."
        
        MsgBox message, vbExclamation, "Health Check Complete"
    End If
    
    ' Also log to immediate window
    Debug.Print "Health check complete for " & sheetName & " - " & issues.Count & " issues found"
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
