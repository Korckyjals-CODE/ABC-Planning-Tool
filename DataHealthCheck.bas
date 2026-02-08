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
' Data Structures for Issue Tracking
' ===========================
' Note: Type definitions are now in HealthCheckTypes.bas module

' ===========================
' Main Health Check Functions
' ===========================

Public Sub RunBasicHealthCheck()
    ' Main entry point for health check
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If IsWeeklyGradebook(ws) Then
        ' Clear previous entries for individual health check runs
        ClearPreviousHealthReportEntries
        ValidateWeeklyGradebookBasic ws
    Else
        MsgBox "Please select a weekly gradebook sheet to run health check.", vbExclamation, "Health Check"
    End If
End Sub

Public Function RunHealthCheckOnWorkbook(ByVal wb As Workbook, Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True) As IssueSummary
    ' Run health check on a specific workbook (for external gradebooks)
    ' Returns aggregated IssueSummary for the entire workbook
    Dim ws As Worksheet
    Dim issuesFound As Boolean
    Dim totalSheets As Long
    Dim processedSheets As Long
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEvents As Boolean
    Dim sheetResult As IssueSummary
    Dim workbookResult As IssueSummary
    Dim TotalIssues As Long
    Dim allIssues As Collection
    
    Set allIssues = New Collection
    TotalIssues = 0
    issuesFound = False
    
    ' Store original Excel settings for performance optimization
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEvents = Application.EnableEvents
    
    ' Disable Excel features for faster execution and to prevent workbook visibility
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo RestoreSettings
    
    ' Clear previous entries only if requested (for new health check runs)
    If ClearReport Then
        ClearPreviousHealthReportEntries
    End If
    
    ' Count total gradebook sheets for progress tracking
    For Each ws In wb.Worksheets
        If IsWeeklyGradebook(ws) Then
            totalSheets = totalSheets + 1
        End If
    Next ws
    
    ' Initialize ProgressBar if we have sheets to process
    Dim MyProgressbar As ProgressBar
    If totalSheets > 0 Then
        Set MyProgressbar = New ProgressBar
        
        With MyProgressbar
            .Title = "Health Check - " & wb.Name
            .ExcelStatusBar = True
            .StartColour = rgbMediumSeaGreen
            .EndColour = rgbGreen
            .TotalActions = totalSheets
        End With
        
        ' Position the subprocess progress bar below the parent (if it exists)
        ' Set manual positioning to override the default centering
        MyProgressbar.StartUpPosition = 0
        
        ' Calculate position below the parent progress bar
        ' Parent is typically centered, so we position this one below it
        Dim parentTop As Long
        parentTop = Application.Top + (Application.Height / 2) - (MyProgressbar.Height / 2)
        
        ' Position this progress bar below the parent with a margin
        MyProgressbar.Left = Application.Left + (Application.Width / 2) - (MyProgressbar.Width / 2)
        MyProgressbar.Top = parentTop + MyProgressbar.Height + 20  ' 20 pixels margin below parent
        
        MyProgressbar.ShowBar
    End If
    
    ' Process worksheets with progress tracking
    For Each ws In wb.Worksheets
        If IsWeeklyGradebook(ws) Then
            processedSheets = processedSheets + 1
            MyProgressbar.NextAction "Checking '" & ws.Name & "'", True
            Debug.Print "Running health check on: " & wb.Name & " - " & ws.Name
            sheetResult = ValidateWeeklyGradebookBasic(ws, MyProgressbar, SuppressDialogs)
            
            ' Aggregate results
            TotalIssues = TotalIssues + sheetResult.IssueCount
            
            ' Add individual issues to aggregated collection
            Dim j As Long
            For j = 1 To sheetResult.Issues.count
                allIssues.Add sheetResult.Issues(j)
            Next j
            
            issuesFound = True
        End If
    Next ws
    
    ' Complete progress bar
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Complete 1, "Health check complete for " & wb.Name
    End If
    
    If Not issuesFound And Not SuppressDialogs Then
        MsgBox "No weekly gradebook sheets found in: " & wb.Name, vbInformation, "Health Check"
    End If
    
    ' Create aggregated workbook result
    workbookResult.filePath = wb.FullName
    workbookResult.fileName = wb.Name
    workbookResult.sheetName = ""  ' Aggregated result
    workbookResult.WorkbookName = wb.Name
    workbookResult.IssueCount = TotalIssues
    Set workbookResult.Issues = allIssues
    
    RunHealthCheckOnWorkbook = workbookResult
    
RestoreSettings:
    ' Restore Excel settings
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEvents
    
    If Err.Number <> 0 Then
        Debug.Print "Error in RunHealthCheckOnWorkbook: " & Err.Description
        Err.Clear
    End If
End Function

Public Function RunHealthCheckOnFile(ByVal filePath As String, Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True, Optional ByVal xlApp As Object = Nothing) As IssueSummary
    ' Run health check on a specific file path
    ' If xlApp is provided, use it; otherwise create a new instance
    ' Returns aggregated IssueSummary for the file
    Dim wb As Workbook
    Dim xlAppCreated As Boolean
    Dim originalVisibleState As Boolean
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEvents As Boolean
    Dim sheetResults As Collection
    Dim sheetResult As IssueSummary
    Dim fileResult As IssueSummary
    Dim fso As Object
    Dim i As Long
    Dim TotalIssues As Long
    Dim allIssues As Collection
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set allIssues = New Collection
    xlAppCreated = False
    
    ' Use provided Excel instance or create a new one
    If xlApp Is Nothing Then
        ' Try to get existing Excel instance first
        On Error Resume Next
        Set xlApp = GetObject(, "Excel.Application")
        If Err.Number <> 0 Then
            Set xlApp = CreateObject("Excel.Application")
            xlAppCreated = True
        End If
        On Error GoTo 0
        
        If xlApp Is Nothing Then
            If Not SuppressDialogs Then
                MsgBox "Could not access Excel application to open file: " & filePath, vbExclamation, "Health Check"
            End If
            ' Return empty result
            fileResult.filePath = filePath
            fileResult.fileName = fso.GetFileName(filePath)
            fileResult.sheetName = ""
            fileResult.WorkbookName = fso.GetBaseName(filePath)
            fileResult.IssueCount = 0
            Set fileResult.Issues = New Collection
            RunHealthCheckOnFile = fileResult
            Exit Function
        End If
    End If
    
    ' Store original settings and ensure Excel instance is completely hidden
    On Error Resume Next
    originalVisibleState = xlApp.Visible
    originalScreenUpdating = xlApp.ScreenUpdating
    originalCalculation = xlApp.Calculation
    originalEvents = xlApp.EnableEvents
    On Error GoTo 0
    
    ' Configure Excel instance for invisible operation with error handling
    On Error Resume Next
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    ' Skip Calculation property setting to avoid runtime error
    ' xlApp.Calculation = -4135  ' This can cause runtime error on some systems
    xlApp.EnableEvents = False
    xlApp.DisplayAlerts = False
    On Error GoTo 0
    
    ' Open the workbook with hidden state
    On Error Resume Next
    Set wb = xlApp.Workbooks.Open(filePath, UpdateLinks:=False, ReadOnly:=True)
    If Err.Number <> 0 Then
        If Not SuppressDialogs Then
            MsgBox "Could not open file: " & filePath & vbCrLf & "Error: " & Err.Description, vbExclamation, "Health Check"
        End If
        Err.Clear
        ' Restore settings before cleanup
        On Error Resume Next
        xlApp.Visible = originalVisibleState
        xlApp.ScreenUpdating = originalScreenUpdating
        ' Skip Calculation property restoration since we didn't set it
        ' xlApp.Calculation = originalCalculation
        xlApp.EnableEvents = originalEvents
        xlApp.DisplayAlerts = True
        On Error GoTo 0
        If xlAppCreated Then xlApp.Quit
        ' Return empty result
        fileResult.filePath = filePath
        fileResult.fileName = fso.GetFileName(filePath)
        fileResult.sheetName = ""
        fileResult.WorkbookName = fso.GetBaseName(filePath)
        fileResult.IssueCount = 0
        Set fileResult.Issues = New Collection
        RunHealthCheckOnFile = fileResult
        Exit Function
    End If
    On Error GoTo 0
    
    ' Ensure all workbook windows are completely hidden
    Dim w As Window
    For Each w In wb.Windows
        w.Visible = False
    Next w
    
    ' Run health check and get results
    sheetResult = RunHealthCheckOnWorkbook(wb, SuppressDialogs, ClearReport)
    
    ' Use the aggregated workbook result directly
    TotalIssues = sheetResult.IssueCount
    Set allIssues = sheetResult.Issues
    
    ' Create aggregated file result
    fileResult.filePath = filePath
    fileResult.fileName = fso.GetFileName(filePath)
    fileResult.sheetName = ""  ' Aggregated result, no single sheet
    fileResult.WorkbookName = fso.GetBaseName(filePath)
    fileResult.IssueCount = TotalIssues
    Set fileResult.Issues = allIssues
    
    ' Close workbook
    wb.Close SaveChanges:=False
    
    ' Restore Excel instance settings
    On Error Resume Next
    xlApp.Visible = originalVisibleState
    xlApp.ScreenUpdating = originalScreenUpdating
    ' Skip Calculation property restoration since we didn't set it
    ' xlApp.Calculation = originalCalculation
    xlApp.EnableEvents = originalEvents
    xlApp.DisplayAlerts = True
    On Error GoTo 0
    
    ' Only quit if we created the instance
    If xlAppCreated Then xlApp.Quit
    
    RunHealthCheckOnFile = fileResult
End Function

Public Sub RunHealthCheckOnSheet(ByVal ws As Worksheet)
    ' Run health check on specific worksheet
    If IsWeeklyGradebook(ws) Then
        Dim result As IssueSummary
        result = ValidateWeeklyGradebookBasic(ws)
    Else
        MsgBox "Selected sheet is not a recognized weekly gradebook format.", vbExclamation, "Health Check"
    End If
End Sub

' ===========================
' Backward Compatibility Wrappers
' ===========================

Public Sub RunHealthCheckOnWorkbookSub(ByVal wb As Workbook, Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True)
    ' Backward compatibility wrapper for RunHealthCheckOnWorkbook
    Dim result As IssueSummary
    result = RunHealthCheckOnWorkbook(wb, SuppressDialogs, ClearReport)
End Sub

Public Sub RunHealthCheckOnFileSub(ByVal filePath As String, Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True, Optional ByVal xlApp As Object = Nothing)
    ' Backward compatibility wrapper for RunHealthCheckOnFile
    Dim result As IssueSummary
    result = RunHealthCheckOnFile(filePath, SuppressDialogs, ClearReport, xlApp)
End Sub

Public Sub RunHealthCheckOnFolderSub(ByVal folderPath As String, Optional ByVal bimester As String = "", Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True)
    ' Backward compatibility wrapper for RunHealthCheckOnFolder
    Dim result As FolderSummary
    result = RunHealthCheckOnFolder(folderPath, bimester, SuppressDialogs, ClearReport)
End Sub

Public Sub RunHealthCheckOnBimesterSub(ByVal bimesterFolderPath As String)
    ' Backward compatibility wrapper for RunHealthCheckOnBimester
    Dim result As BimesterSummary
    result = RunHealthCheckOnBimester(bimesterFolderPath)
End Sub

' ===========================
' Weekly Gradebook Detection
' ===========================

Private Function IsWeeklyGradebook(ByVal ws As Worksheet) As Boolean
    ' Detect if worksheet is a weekly gradebook based on structure
    Dim hasWeeklyStructure As Boolean
    hasWeeklyStructure = False
    
    ' Check for common weekly gradebook indicators
    If ws.Cells(1, 1).value Like "*Nota Semanal*" Or _
       ws.Cells(1, 3).value Like "*Nota Semanal*" Then
        hasWeeklyStructure = True
    End If
    
    ' Check for class columns (Clase 1, Clase 2, etc.)
    If Not hasWeeklyStructure Then
        Dim col As Integer
        For col = 3 To 7 ' Check columns C through G
            If ws.Cells(3, col).value Like "Clase *" Then
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
            If IsNumeric(ws.Cells(2, col).value) And ws.Cells(2, col).value Like "*%" Then
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

Private Function ValidateWeeklyGradebookBasic(ByVal ws As Worksheet, Optional ByRef MyProgressbar As ProgressBar = Nothing, Optional ByVal SuppressDialogs As Boolean = False) As IssueSummary
    Dim Issues As Collection
    Dim result As IssueSummary
    Dim fso As Object
    
    Set Issues = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Log start of validation
    Debug.Print "Starting health check for: " & ws.Name
    
    ' Check for classes with default grades but non-zero weight
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.StatusMessage = "Checking default grades with weights for '" & ws.Name & "'"
    End If
    CheckDefaultGradesWithWeight ws, Issues
    
    ' Check weight sum
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.StatusMessage = "Checking weight sum for '" & ws.Name & "'"
    End If
    CheckWeightSum ws, Issues
    
    ' Check for empty student rows with default grades
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.StatusMessage = "Checking empty student rows for '" & ws.Name & "'"
    End If
    CheckEmptyStudentRows ws, Issues
    
    ' Report results to health report sheet
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.StatusMessage = "Generating health report for '" & ws.Name & "'"
    End If
    ReportHealthIssuesToSheet Issues, ws.Name, ws.Parent.Name, SuppressDialogs
    
    ' Create and populate IssueSummary
    result.filePath = ws.Parent.FullName
    result.fileName = ws.Parent.Name
    result.sheetName = ws.Name
    result.WorkbookName = ws.Parent.Name
    result.IssueCount = Issues.count
    Set result.Issues = Issues
    
    ValidateWeeklyGradebookBasic = result
End Function

Private Sub CheckDefaultGradesWithWeight(ByVal ws As Worksheet, ByRef Issues As Collection)
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
            Issues.Add issue
            Debug.Print "ISSUE: " & issue
        End If
    Next classCol
End Sub

Private Sub CheckWeightSum(ByVal ws As Worksheet, ByRef Issues As Collection)
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
        Issues.Add issue
        Debug.Print "ISSUE: " & issue
    End If
End Sub

Private Sub CheckEmptyStudentRows(ByVal ws As Worksheet, ByRef Issues As Collection)
    ' Check for student rows that have default grades but no name
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    
    Dim row As Long
    For row = 4 To lastRow ' Assuming student data starts at row 4
        If IsEmptyStudentRow(ws, row) Then
            Dim issue As String
            issue = "Row " & row & " appears to be empty student data with default grades"
            Issues.Add issue
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
        If ws.Cells(3, col).value Like "Clase *" Then
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
        If Not IsEmpty(ws.Cells(row, colNum).value) Then
            If ws.Cells(row, colNum).value <> DEFAULT_GRADE_VALUE Then
                AllGradesEqualDefault = False
                Exit Function
            End If
        End If
    Next row
    
    AllGradesEqualDefault = True
End Function

Private Function GetClassWeight(ByVal ws As Worksheet, ByVal colNum As Integer) As Double
    ' Get the weight percentage for a class column
    ' Handle percentage format cells (decimal values 0-1)
    Dim cellValue As Variant
    cellValue = ws.Cells(2, colNum).value
    
    ' Check if cell contains a numeric value
    If IsNumeric(cellValue) Then
        ' Convert decimal to percentage (0.1 = 10%, 1.0 = 100%)
        GetClassWeight = CDbl(cellValue) * 100
    Else
        ' Fallback for text-based percentages (legacy support)
        Dim weightStr As String
        weightStr = CStr(cellValue)
        If weightStr Like "*%" Then
            GetClassWeight = CDbl(Replace(weightStr, "%", ""))
        Else
            GetClassWeight = 0
        End If
    End If
End Function

Private Function GetClassNumber(ByVal ws As Worksheet, ByVal colNum As Integer) As String
    ' Extract class number from header (e.g., "Clase 1" -> "1")
    Dim header As String
    header = CStr(ws.Cells(3, colNum).value)
    
    If header Like "Clase *" Then
        GetClassNumber = Replace(header, "Clase ", "")
    Else
        GetClassNumber = "Unknown"
    End If
End Function

Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    ' Find the last row with data in column A (student names)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' If no data found, return 0
    If lastRow = 1 And IsEmpty(ws.Cells(1, 1).value) Then
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
    hasName = Not IsEmpty(ws.Cells(row, 1).value) And ws.Cells(row, 1).value <> "0"
    
    ' Check if all class columns have default grades
    hasDefaultGrades = True
    Dim classColumns As Collection
    Set classColumns = GetClassColumns(ws)
    
    Dim classCol As Variant
    For Each classCol In classColumns
        If ws.Cells(row, CInt(classCol)).value <> DEFAULT_GRADE_VALUE Then
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

Private Sub ReportHealthIssuesToSheet(ByVal Issues As Collection, ByVal sheetName As String, ByVal WorkbookName As String, Optional ByVal SuppressDialogs As Boolean = False)
    ' Create or update health report sheet with detailed results
    Dim reportWs As Worksheet
    Dim screenUpdateState As Boolean
    Dim calcState As XlCalculation
    Dim eventsState As Boolean
    
    ' Store current Excel settings for performance optimization
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    ' Disable Excel features for faster execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo RestoreSettings
    
    Set reportWs = GetOrCreateHealthReportSheet
    
    ' Add new health check entry
    AddHealthCheckEntry reportWs, Issues, sheetName, WorkbookName
    
    ' Show summary message only if not suppressed
    If Not SuppressDialogs Then
        If Issues.count = 0 Then
            MsgBox "[OK] " & sheetName & " - No health issues found!" & vbCrLf & vbCrLf & _
                   "The gradebook data appears to be healthy." & vbCrLf & vbCrLf & _
                   "Check the 'HealthReport' sheet for detailed results.", vbInformation, "Health Check Complete"
        Else
            MsgBox "[WARNING] " & sheetName & " - " & Issues.count & " health issue(s) found!" & vbCrLf & vbCrLf & _
                   "Please review the 'HealthReport' sheet for detailed information." & vbCrLf & vbCrLf & _
                   "You can now work with your gradebook while keeping the report open.", vbExclamation, "Health Check Complete"
        End If
    End If
    
    ' Also log to immediate window
    Debug.Print "Health check complete for " & sheetName & " - " & Issues.count & " issues found"
    
    ' Auto-fit columns for better readability
    AutoFitHealthReportColumns reportWs
    
RestoreSettings:
    ' Restore Excel settings
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    
    If Err.Number <> 0 Then
        Debug.Print "Error in ReportHealthIssuesToSheet: " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub AutoFitHealthReportColumns(ByVal reportWs As Worksheet)
    ' Auto-fit all columns in the health report for better readability
    On Error Resume Next
    
    ' Auto-fit columns A through F (all data columns)
    reportWs.Columns("A:F").AutoFit
    
    ' Ensure minimum column widths for better readability
    If reportWs.Columns("A").ColumnWidth < 15 Then reportWs.Columns("A").ColumnWidth = 15
    If reportWs.Columns("B").ColumnWidth < 20 Then reportWs.Columns("B").ColumnWidth = 20
    If reportWs.Columns("C").ColumnWidth < 15 Then reportWs.Columns("C").ColumnWidth = 15
    If reportWs.Columns("D").ColumnWidth < 12 Then reportWs.Columns("D").ColumnWidth = 12
    If reportWs.Columns("E").ColumnWidth < 10 Then reportWs.Columns("E").ColumnWidth = 10
    If reportWs.Columns("F").ColumnWidth < 30 Then reportWs.Columns("F").ColumnWidth = 30
    
    On Error GoTo 0
End Sub

Private Function GetOrCreateHealthReportSheet() As Worksheet
    ' Get existing health report sheet or create a new one in the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook by finding the workbook that contains this code
    Set planningWb = GetPlanningWorkbook
    
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

Private Function GetPlanningWorkbook() As Workbook
    ' Get the planning workbook that contains this VBA code (optimized)
    Dim wb As Workbook
    Dim ws As Worksheet
    
    ' First try: Check if ThisWorkbook has planning characteristics
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("wsSchedule")
    If Not ws Is Nothing Then
        Set GetPlanningWorkbook = ThisWorkbook
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' Second try: Look for workbooks with "Planning" in the name (faster than sheet checking)
    For Each wb In Application.Workbooks
        If wb.Name Like "*Planning*" Then
            Set GetPlanningWorkbook = wb
            Exit Function
        End If
    Next wb
    
    ' Third try: Check for characteristic sheets (only if name search fails)
    For Each wb In Application.Workbooks
        On Error Resume Next
        Set ws = wb.Worksheets("wsSchedule")
        If Not ws Is Nothing Then
            Set GetPlanningWorkbook = wb
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next wb
    
    ' Fallback: Use ThisWorkbook
    Set GetPlanningWorkbook = ThisWorkbook
End Function

Private Sub SetupHealthReportSheet(ByVal ws As Worksheet)
    ' Set up the health report sheet with headers and formatting (optimized)
    Dim headerData As Variant
    Dim instructionData As Variant
    Dim headerRange As Range
    Dim instructionRange As Range
    
    ' Prepare data arrays for bulk operations
    headerData = Array("Data Health Check Report", "Generated on: " & Format(Now, "yyyy-mm-dd hh:mm:ss"), "", "", _
                      "Timestamp", "Workbook", "Sheet", "Status", "Issues Count", "Issue Details")
    
    instructionData = Array("Instructions:", "- Green status = No issues found", _
                           "- Red status = Issues found - review details in column F", _
                           "- Click on any row to see more details", _
                           "- This report updates automatically with each health check")
    
    With ws
        ' Clear any existing content
        .Cells.Clear
        
        ' Bulk write header data
        .Range("A1:A10").value = Application.Transpose(headerData)
        
        ' Format main header (A1)
        With .Cells(1, 1)
            .Font.Bold = True
            .Font.Size = 16
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Format subtitle (A2)
        .Cells(2, 1).Font.Italic = True
        
        ' Format column headers (A4:F4)
        Set headerRange = .Range("A4:F4")
        With headerRange
            .Font.Bold = True
            .Interior.Color = RGB(217, 225, 242)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
        
        ' Set column widths in one operation
        .Range("A:A").ColumnWidth = 20
        .Range("B:B").ColumnWidth = 25
        .Range("C:C").ColumnWidth = 20
        .Range("D:D").ColumnWidth = 15
        .Range("E:E").ColumnWidth = 12
        .Range("F:F").ColumnWidth = 60
        
        ' Bulk write instruction data
        .Range("A6:A10").value = Application.Transpose(instructionData)
        
        ' Format instruction text
        Set instructionRange = .Range("A7:A10")
        With instructionRange
            .Font.Size = 10
            .Font.Color = RGB(100, 100, 100)
        End With
        
        ' Freeze panes without using Select method
        ' Note: FreezePanes is set relative to the active cell, so we activate the sheet first
        Dim originalActiveSheet As Worksheet
        Set originalActiveSheet = ActiveSheet
        .Activate
        .Range("A5").Activate
        ActiveWindow.FreezePanes = True
        ' Restore original active sheet if different
        If Not originalActiveSheet Is Nothing And originalActiveSheet.Name <> .Name Then
            originalActiveSheet.Activate
        End If
    End With
End Sub

Private Sub AddHealthCheckEntry(ByVal reportWs As Worksheet, ByVal Issues As Collection, ByVal sheetName As String, ByVal WorkbookName As String)
    ' Add a new health check entry to the report sheet (optimized)
    Dim lastRow As Long
    Dim newRow As Long
    Dim dataArray As Variant
    Dim dataRange As Range
    Dim issueDetails As String
    Dim i As Integer
    
    lastRow = reportWs.Cells(reportWs.Rows.count, 1).End(xlUp).row
    
    ' Find next available row (skip instruction rows)
    If lastRow < 10 Then lastRow = 10
    
    newRow = lastRow + 1
    
    ' Prepare data array for bulk write
    ReDim dataArray(1 To 1, 1 To 6)
    dataArray(1, 1) = Format(Now, "yyyy-mm-dd hh:mm:ss")
    dataArray(1, 2) = WorkbookName
    dataArray(1, 3) = sheetName
    dataArray(1, 5) = Issues.count
    
    ' Add status and issue details
    If Issues.count = 0 Then
        dataArray(1, 4) = "[OK] HEALTHY"
        dataArray(1, 6) = "No issues detected"
    Else
        dataArray(1, 4) = "[WARNING] ISSUES"
        
        ' Build issue details string efficiently
        issueDetails = ""
        For i = 1 To Issues.count
            issueDetails = issueDetails & i & ". " & Issues(i)
            If i < Issues.count Then issueDetails = issueDetails & vbLf
        Next i
        dataArray(1, 6) = issueDetails
    End If
    
    ' Bulk write data to worksheet
    Set dataRange = reportWs.Range("A" & newRow & ":F" & newRow)
    dataRange.value = dataArray
    
    ' Apply formatting
    With dataRange
        ' Add borders
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        
        ' Format status column
        If Issues.count = 0 Then
            .Cells(1, 4).Interior.Color = RGB(198, 239, 206)
            .Cells(1, 4).Font.Color = RGB(0, 100, 0)
        Else
            .Cells(1, 4).Interior.Color = RGB(255, 199, 206)
            .Cells(1, 4).Font.Color = RGB(156, 0, 6)
            .Cells(1, 6).WrapText = True
        End If
    End With
    
    ' Auto-fit row height for issue details
    reportWs.Rows(newRow).AutoFit
    
    ' Scroll to the new entry without using Select method
    ' Note: This functionality is preserved by ensuring the sheet is properly formatted
    ' The user can manually scroll to see the new entry if needed
End Sub

' ===========================
' Legacy Reporting Functions (for backward compatibility)
' ===========================

Private Sub ReportHealthIssues(ByVal Issues As Collection, ByVal sheetName As String)
    ' Legacy function - now redirects to sheet-based reporting
    Dim planningWb As Workbook
    Set planningWb = GetPlanningWorkbook
    ReportHealthIssuesToSheet Issues, sheetName, planningWb.Name, False  ' Show dialogs for legacy calls
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
        
        Debug.Print "Class Columns Found: " & classColumns.count
        Dim classCol As Variant
        For Each classCol In classColumns
            Debug.Print "  Column " & classCol & ": " & ws.Cells(3, CInt(classCol)).value & " (Weight: " & GetClassWeight(ws, CInt(classCol)) & "%)"
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
    ' Note: This function does NOT clear previous entries as it's part of a batch process
    ' The calling function (GenerateRawGradebooks) handles clearing previous entries
    
    If wb Is Nothing Then
        Debug.Print "Health Check: Cannot run on Nothing workbook - " & templatePath
        Exit Sub
    End If
    
    ' Check if workbook is still accessible
    On Error Resume Next
    Dim testCount As Long
    testCount = wb.Worksheets.count
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
        ValidateWeeklyGradebookBasic ws, , True  ' Suppress dialogs for generated gradebooks
    Else
        Debug.Print "Health Check: Generated gradebook is not recognized format - " & templatePath
    End If
End Sub

Public Function RunHealthCheckOnFolder(ByVal folderPath As String, Optional ByVal bimester As String = "", Optional ByVal SuppressDialogs As Boolean = False, Optional ByVal ClearReport As Boolean = True) As FolderSummary
    ' Run health check on all gradebook files in a folder
    ' Uses a single hidden Excel instance to prevent flashing and improve performance
    ' Returns FolderSummary with aggregated results
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim processedCount As Integer
    Dim totalFiles As Long
    Dim xlApp As Object
    Dim xlAppCreated As Boolean
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEvents As Boolean
    Dim ext As String
    Dim folderResult As FolderSummary
    Dim fileResult As IssueSummary
    Dim TotalIssues As Long
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    TotalIssues = 0
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbExclamation, "Health Check"
        ' Return empty result
        folderResult.folderPath = folderPath
        folderResult.FolderName = fso.GetBaseName(folderPath)
        folderResult.totalFiles = 0
        folderResult.TotalIssues = 0
        Set folderResult.FileSummaries = Nothing
        RunHealthCheckOnFolder = folderResult
        Exit Function
    End If
    
    ' Clear previous entries only if requested (for new health check runs)
    If ClearReport Then
        ClearPreviousHealthReportEntries
    End If
    
    Set folder = fso.GetFolder(folderPath)
    processedCount = 0
    
    ' Count total .xlsx and .xlsm files for progress tracking
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "xlsx" Or ext = "xlsm" Then
            If bimester = "" Or InStr(1, file.Name, bimester, vbTextCompare) > 0 Then
                totalFiles = totalFiles + 1
            End If
        End If
    Next file
    
    ' Initialize ProgressBar if we have files to process
    Dim MyProgressbar As ProgressBar
    If totalFiles > 0 Then
        Set MyProgressbar = New ProgressBar
        
        With MyProgressbar
            .Title = "Health Check - Folder: " & fso.GetBaseName(folderPath)
            If bimester <> "" Then .Title = .Title & " (" & bimester & ")"
            .ExcelStatusBar = True
            .StartColour = rgbMediumSeaGreen
            .EndColour = rgbGreen
            .TotalActions = totalFiles
        End With
        
        MyProgressbar.ShowBar
    End If
    
    ' Create a single hidden Excel instance for all file processing
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        MsgBox "Could not create Excel instance for health check: " & Err.Description, vbExclamation, "Health Check"
        Err.Clear
        ' Return empty result
        folderResult.folderPath = folderPath
        folderResult.FolderName = fso.GetBaseName(folderPath)
        folderResult.totalFiles = 0
        folderResult.TotalIssues = 0
        Set folderResult.FileSummaries = Nothing
        RunHealthCheckOnFolder = folderResult
        Exit Function
    End If
    On Error GoTo 0
    
    xlAppCreated = True
    
    ' Store original Excel settings for performance optimization
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEvents = Application.EnableEvents
    
    ' Disable Excel features for faster execution and to prevent visibility
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Configure the hidden Excel instance for optimal performance with error handling
    On Error Resume Next
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    ' Skip Calculation property setting to avoid runtime error
    ' xlApp.Calculation = -4135  ' This can cause runtime error on some systems
    xlApp.EnableEvents = False
    xlApp.DisplayAlerts = False
    On Error GoTo 0
    
    Debug.Print "=== Starting Health Check on Folder ==="
    Debug.Print "Folder: " & folderPath
    If bimester <> "" Then Debug.Print "Bimester: " & bimester
    Debug.Print "Files to process: " & totalFiles
    Debug.Print "Using single hidden Excel instance for all files"
    Debug.Print "======================================="
    
    ' Process all files using the single hidden Excel instance
    For Each file In folder.Files
        ext = LCase(fso.GetExtensionName(file.Name))
        If ext = "xlsx" Or ext = "xlsm" Then
            If bimester = "" Or InStr(1, file.Name, bimester, vbTextCompare) > 0 Then
                processedCount = processedCount + 1
                If Not MyProgressbar Is Nothing Then
                    MyProgressbar.NextAction "Processing '" & file.Name & "'", True
                End If
                Debug.Print "Processing: " & file.Name
                fileResult = RunHealthCheckOnFile(file.path, True, False, xlApp)  ' Use the shared hidden Excel instance
                TotalIssues = TotalIssues + fileResult.IssueCount
            End If
        End If
    Next file
    
    ' Complete progress bar
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Complete 1, "Health check complete on " & processedCount & " files"
    End If
    
    ' Clean up the hidden Excel instance
    If Not xlApp Is Nothing Then
        xlApp.Quit
        Set xlApp = Nothing
    End If
    
    ' Restore original Excel settings
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEvents
    
    Debug.Print "=== Health Check Complete ==="
    Debug.Print "Files processed: " & processedCount
    Debug.Print "============================="
    
    ' Create folder result
    folderResult.folderPath = folderPath
    folderResult.FolderName = fso.GetBaseName(folderPath)
    folderResult.totalFiles = processedCount
    folderResult.TotalIssues = TotalIssues
    Set folderResult.FileSummaries = Nothing  ' Simplified - no detailed file tracking
    
    ' Show summary dialog only if not suppressed
    If Not SuppressDialogs Then
        MsgBox "Health check completed on " & processedCount & " files in folder: " & folderPath, vbInformation, "Health Check Complete"
    End If
    
    RunHealthCheckOnFolder = folderResult
End Function

Public Function RunHealthCheckOnBimester(ByVal bimesterFolderPath As String) As BimesterSummary
    ' Run health check on all week subfolders in a bimester folder
    ' Each week subfolder (W01, W02, etc.) will be processed by RunHealthCheckOnFolder
    ' Uses three-level progress bars: Main (bimester) -> Folder (week) -> File (individual files)
    ' Returns BimesterSummary with aggregated results
    Dim fso As Object
    Dim bimesterFolder As Object
    Dim weekFolder As Object
    Dim processedFolders As Integer
    Dim TotalFolders As Long
    Dim totalFiles As Long
    Dim TotalIssues As Long
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEvents As Boolean
    Dim weekFolderName As String
    Dim weekFolderPath As String
    Dim bimesterResult As BimesterSummary
    Dim folderResult As FolderSummary
    Dim i As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    totalFiles = 0
    TotalIssues = 0
    
    If Not fso.FolderExists(bimesterFolderPath) Then
        MsgBox "Bimester folder not found: " & bimesterFolderPath, vbExclamation, "Health Check"
        ' Return empty result
        bimesterResult.BimesterPath = bimesterFolderPath
        bimesterResult.BimesterName = fso.GetBaseName(bimesterFolderPath)
        bimesterResult.TotalFolders = 0
        bimesterResult.totalFiles = 0
        bimesterResult.TotalIssues = 0
        Set bimesterResult.FolderSummaries = Nothing
        RunHealthCheckOnBimester = bimesterResult
        Exit Function
    End If
    
    ' Clear previous entries for bimester health check runs
    ClearPreviousHealthReportEntries
    
    Set bimesterFolder = fso.GetFolder(bimesterFolderPath)
    processedFolders = 0
    
    ' Count total week folders (WXX pattern) for progress tracking
    For Each weekFolder In bimesterFolder.SubFolders
        weekFolderName = weekFolder.Name
        If weekFolderName Like "W##" Then  ' Pattern: W followed by exactly 2 digits
            TotalFolders = TotalFolders + 1
        End If
    Next weekFolder
    
    ' Store original Excel settings for performance optimization
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEvents = Application.EnableEvents
    
    ' Disable Excel features for faster execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo RestoreSettings
    
    ' Initialize Main ProgressBar (top level - bimester)
    Dim MainProgressbar As ProgressBar
    If TotalFolders > 0 Then
        Set MainProgressbar = New ProgressBar
        
        With MainProgressbar
            .Title = "Health Check - Bimester: " & fso.GetBaseName(bimesterFolderPath)
            .ExcelStatusBar = True
            .StartColour = rgbMediumSeaGreen
            .EndColour = rgbGreen
            .TotalActions = TotalFolders
        End With
        
        ' Position the main progress bar at the top
        MainProgressbar.StartUpPosition = 0
        MainProgressbar.Left = Application.Left + (Application.Width / 2) - (MainProgressbar.Width / 2)
        MainProgressbar.Top = Application.Top + (Application.Height / 2) - (MainProgressbar.Height / 2) - 100  ' 100 pixels above center
        
        MainProgressbar.ShowBar
    End If
    
    Debug.Print "=== Starting Health Check on Bimester ==="
    Debug.Print "Bimester Folder: " & bimesterFolderPath
    Debug.Print "Week folders to process: " & TotalFolders
    Debug.Print "========================================="
    
    ' Process all week folders
    For Each weekFolder In bimesterFolder.SubFolders
        weekFolderName = weekFolder.Name
        weekFolderPath = weekFolder.path
        
        If weekFolderName Like "W##" Then  ' Pattern: W followed by exactly 2 digits
            processedFolders = processedFolders + 1
            If Not MainProgressbar Is Nothing Then
                MainProgressbar.NextAction "Processing week folder '" & weekFolderName & "'", True
            End If
            
            Debug.Print "Processing week folder: " & weekFolderName
            
            ' Run health check on this week folder (suppress its summary dialog and don't clear report)
            folderResult = RunHealthCheckOnFolder(weekFolderPath, "", True, False)  ' Suppress dialogs and don't clear report for subprocess
            
            ' Aggregate totals
            totalFiles = totalFiles + folderResult.totalFiles
            TotalIssues = TotalIssues + folderResult.TotalIssues
            
            Debug.Print "  - Files processed: " & folderResult.totalFiles
            Debug.Print "  - Issues found: " & folderResult.TotalIssues
        End If
    Next weekFolder
    
    ' Complete main progress bar
    If Not MainProgressbar Is Nothing Then
        MainProgressbar.Complete 1, "Bimester health check complete"
    End If
    
    Debug.Print "=== Bimester Health Check Complete ==="
    Debug.Print "Week folders processed: " & processedFolders
    Debug.Print "Total files processed: " & totalFiles
    Debug.Print "Total issues found: " & TotalIssues
    Debug.Print "====================================="
    
    ' Create bimester result
    bimesterResult.BimesterPath = bimesterFolderPath
    bimesterResult.BimesterName = fso.GetBaseName(bimesterFolderPath)
    bimesterResult.TotalFolders = processedFolders
    bimesterResult.totalFiles = totalFiles
    bimesterResult.TotalIssues = TotalIssues
    Set bimesterResult.FolderSummaries = Nothing  ' Simplified - no detailed folder tracking
    
    ' Show final summary dialog
    Dim summaryMessage As String
    summaryMessage = "Bimester Health Check Complete!" & vbCrLf & vbCrLf & _
                    "Week folders processed: " & processedFolders & vbCrLf & _
                    "Total files processed: " & totalFiles & vbCrLf & _
                    "Total issues found: " & TotalIssues & vbCrLf & vbCrLf & _
                    "Check the 'HealthReport' sheet for detailed results."
    
    If TotalIssues = 0 Then
        MsgBox summaryMessage, vbInformation, "Health Check Complete - All Good!"
    Else
        MsgBox summaryMessage, vbExclamation, "Health Check Complete - Issues Found"
    End If
    
    RunHealthCheckOnBimester = bimesterResult
    
RestoreSettings:
    ' Restore Excel settings
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEvents
    
    If Err.Number <> 0 Then
        Debug.Print "Error in RunHealthCheckOnBimester: " & Err.Description
        Err.Clear
    End If
End Function

' ===========================
' Helper Functions for Bimester Health Check
' ===========================

Private Function CountExcelFilesInFolder(ByVal folderPath As String) As Long
    ' Count the number of Excel files (.xlsx and .xlsm) in a folder
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ext As String
    Dim count As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    count = 0
    
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        For Each file In folder.Files
            ext = LCase(fso.GetExtensionName(file.Name))
            If ext = "xlsx" Or ext = "xlsm" Then
                count = count + 1
            End If
        Next file
    End If
    
    CountExcelFilesInFolder = count
End Function

Private Function CountIssuesInHealthReport(ByVal weekFolderName As String) As Long
    ' Count the number of issues found for a specific week folder in the health report
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    Dim lastRow As Long
    Dim row As Long
    Dim count As Long
    Dim sheetName As String
    
    count = 0
    
    ' Get the planning workbook and health report sheet
    Set planningWb = GetPlanningWorkbook
    On Error Resume Next
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    If Not reportWs Is Nothing Then
        lastRow = reportWs.Cells(reportWs.Rows.count, 1).End(xlUp).row
        
        ' Count issues for this week folder (check if sheet name contains week folder name)
        For row = 11 To lastRow  ' Data starts at row 11
            sheetName = CStr(reportWs.Cells(row, 3).value)  ' Sheet name is in column C
            If InStr(1, sheetName, weekFolderName, vbTextCompare) > 0 Then
                ' Check if this row has issues (Issues Count > 0)
                If IsNumeric(reportWs.Cells(row, 5).value) Then
                    count = count + CLng(reportWs.Cells(row, 5).value)
                End If
            End If
        Next row
    End If
    
    CountIssuesInHealthReport = count
End Function

' ===========================
' Health Report Management
' ===========================

Public Sub ClearHealthReport()
    ' Clear the health report sheet from the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = GetPlanningWorkbook
    
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

Public Sub ClearPreviousHealthReportEntries()
    ' Clear previous health report entries but keep the sheet structure
    ' This is used when starting a new health check run
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = GetPlanningWorkbook
    
    On Error Resume Next
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    If Not reportWs Is Nothing Then
        ' Clear all data rows but keep the header and instruction rows (rows 1-10)
        Dim lastRow As Long
        lastRow = reportWs.Cells(reportWs.Rows.count, 1).End(xlUp).row
        
        If lastRow > 10 Then
            ' Clear data rows (from row 11 onwards)
            reportWs.Range("A11:F" & lastRow).ClearContents
            Debug.Print "Cleared previous health report entries (rows 11-" & lastRow & ")"
        End If
        
        ' Update the generation timestamp
        reportWs.Cells(2, 1).value = "Generated on: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        
        ' Regenerate the instruction section with corrected characters
        SetupHealthReportSheet reportWs
    End If
End Sub

Public Sub ShowHealthReport()
    ' Show the health report sheet from the planning workbook
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    
    ' Get the planning workbook (where this code resides)
    Set planningWb = GetPlanningWorkbook
    
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

' ===========================
' Health Check Validation for GenerateRawGradebooks
' ===========================

Public Function CheckHealthReportForIssues(Optional ByVal hoursThreshold As Integer = 24) As Boolean
    ' Check if HealthReport sheet has any unresolved issues
    ' Returns True if issues found, False if no issues and recent health check exists
    ' Parameters:
    '   hoursThreshold: How many hours back to consider a health check "recent" (default: 24)
    
    Dim reportWs As Worksheet
    Dim planningWb As Workbook
    Dim lastRow As Long
    Dim i As Long
    Dim hasRecentHealthCheck As Boolean
    Dim hasIssues As Boolean
    Dim healthCheckThreshold As Date
    
    ' Define threshold: health check should be within specified hours
    healthCheckThreshold = Now - (hoursThreshold / 24)
    
    ' Get the health report sheet
    On Error Resume Next
    Set planningWb = GetPlanningWorkbook
    Set reportWs = planningWb.Worksheets(HEALTH_REPORT_SHEET_NAME)
    On Error GoTo 0
    
    ' If no health report exists, return True (assume issues exist)
    If reportWs Is Nothing Then
        CheckHealthReportForIssues = True
        Debug.Print "Health Check Validation: No health report found - assuming issues exist"
        Exit Function
    End If
    
    lastRow = reportWs.Cells(reportWs.Rows.count, 1).End(xlUp).row
    hasRecentHealthCheck = False
    hasIssues = False
    
    ' Check health report entries (skip header rows 1-10)
    For i = 11 To lastRow
        Dim timestamp As Date
        Dim status As String
        Dim IssueCount As Long
        
        On Error Resume Next
        timestamp = reportWs.Cells(i, 1).value
        status = reportWs.Cells(i, 4).value
        IssueCount = reportWs.Cells(i, 5).value
        On Error GoTo 0
        
        ' Check if this is a recent health check (within threshold)
        If timestamp > healthCheckThreshold Then
            hasRecentHealthCheck = True
            
            ' Check if there are issues
            If IssueCount > 0 Or InStr(status, "ISSUES") > 0 Then
                hasIssues = True
                Debug.Print "Health Check Validation: Issues found in recent health check - " & reportWs.Cells(i, 2).value & " / " & reportWs.Cells(i, 3).value
                Exit For ' Found issues, no need to continue
            End If
        End If
    Next i
    
    ' Log the validation result
    If Not hasRecentHealthCheck Then
        Debug.Print "Health Check Validation: No recent health check found (within " & hoursThreshold & " hours)"
    ElseIf Not hasIssues Then
        Debug.Print "Health Check Validation: Recent health check found with no issues"
    End If
    
    ' Return True if issues found OR no recent health check
    CheckHealthReportForIssues = (hasIssues Or Not hasRecentHealthCheck)
End Function

' ===========================
' Test Function for Health Check Validation
' ===========================

Public Sub TestHealthCheckValidation()
    ' Test function to verify the health check validation works correctly
    ' This can be called from the immediate window for testing
    
    Debug.Print "=== Testing Health Check Validation ==="
    
    ' Test with different time thresholds
    Dim hasIssues24h As Boolean
    Dim hasIssues1h As Boolean
    Dim hasIssues168h As Boolean ' 1 week
    
    hasIssues24h = CheckHealthReportForIssues(24)
    hasIssues1h = CheckHealthReportForIssues(1)
    hasIssues168h = CheckHealthReportForIssues(168)
    
    Debug.Print "Health Check Status (24h threshold): " & IIf(hasIssues24h, "ISSUES FOUND", "NO ISSUES")
    Debug.Print "Health Check Status (1h threshold): " & IIf(hasIssues1h, "ISSUES FOUND", "NO ISSUES")
    Debug.Print "Health Check Status (168h threshold): " & IIf(hasIssues168h, "ISSUES FOUND", "NO ISSUES")
    
    Debug.Print "=== Test Complete ==="
End Sub
