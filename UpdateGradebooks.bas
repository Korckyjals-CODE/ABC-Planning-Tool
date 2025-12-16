Attribute VB_Name = "UpdateGradebooks"
Option Explicit

' Constants
Private Const FORMULA_START_CELL As String = "C5"

' ===========================
' GenerateRawGradebooks
' ===========================
' Orchestrates:
' 1) Empty temp folder
' 2) Copy source -> temp
' 3) Set bimester subfolder path
' 4) For each template .xlsx in bimester folder:
'    - Parse grade tag ? grade code
'    - Filter by grade levels if provided (optional)
'    - Open matching grade workbook(s) from each subfolder (so formulas can link to open files)
'    - Open template (use open instance if already open)
'    - Place grade lookup formula in C5 and copy to rectangular range (C5 : lastRow, lastCol)
'    - Convert formulas to values in rectangular range (C5 : lastRow, lastCol) per rules
'    - Clear zero grade values in weekly columns
'    - Place integrative activity formula in "Actividad Integradora" column
'    - Convert integrative activity formulas to values
'    - Clear zero grade values in integrative activity column
'    - Save & close template
'    - Close only the subfolder files we opened for this template
'
' Parameters:
' - strBimester: The bimester folder name (e.g., "B1", "B2")
' - gradeLevels: Optional. Can be a single string or array of strings containing grade levels to process.
'   Supported grade levels: "1A", "1B", "2A", "2B", ..., "12A", "12B", "DC3A", "DC3B", "DC3BSEC"
'   If not provided, processes all grade levels found in the bimester folder.
'
' Logging:
' - Immediate window (Debug.Print)
' - Sheet "GRB_Log" in ThisWorkbook (cleared/created each run)
'
' Notes:
' - Paths are regular Windows/OneDrive local paths (e.g., C:\Users\...\OneDrive - ...\2526\Computers\Grades)
' - Only immediate subfolders of strBimesterFolderURL are considered for step 4.3
' - Only .xlsx files are processed/opened
' - Calculation set to Manual during run, then restored and recalculated once
' - ScreenUpdating/Alerts handled
'
Public Sub GenerateRawGradebooks(ByVal strBimester As String, Optional ByVal gradeLevels As Variant = Empty)
    Dim strGradebooksTempFolderURL As String
    Dim strSourceFolderURL As String
    Dim strBimesterFolderURL As String
    
    ' ==== PLACEHOLDERS: set these three before running ====
    strGradebooksTempFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Temp_Grades\")
    strSourceFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Grades\")
    strBimesterFolderURL = JoinPath(strGradebooksTempFolderURL, strBimester)   ' e.g., B1

    ' ======================================================
    ' Initialize logging for health check validation
    ' ======================================================
    Dim logLines As Collection
    Set logLines = New Collection

    ' ======================================================
    ' Health Check Validation - NEW
    ' ======================================================
    
    ' Check if health report has unresolved issues
    If CheckHealthReportForIssues(24) Then ' 24 hour threshold
        Dim response As VbMsgBoxResult
        response = MsgBox("WARNING: Health check issues detected or no recent health check found." & vbCrLf & vbCrLf & _
                         "Running GenerateRawGradebooks with files that have issues may lead to incorrect grades." & vbCrLf & vbCrLf & _
                         "Do you want to:" & vbCrLf & _
                         "• YES: View health report and fix issues first" & vbCrLf & _
                         "• NO: Proceed anyway (not recommended)" & vbCrLf & _
                         "• CANCEL: Abort and run health check manually", _
                         vbYesNoCancel + vbExclamation, "Health Check Required")
        
        Select Case response
            Case vbYes
                ' Show health report for user to review
                ShowHealthReport
                MsgBox "Please review the health report, fix all issues, then run GenerateRawGradebooks again.", _
                       vbInformation, "Fix Issues First"
                Exit Sub
                
            Case vbNo
                Log logLines, "WARNING: Proceeding with GenerateRawGradebooks despite health check issues"
                
            Case vbCancel
                Log logLines, "GenerateRawGradebooks aborted by user - health check issues must be resolved"
                Exit Sub
        End Select
    End If

    ' ======================================================
    
    Dim fso As Object
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEvents As Boolean
    Dim booEnablePerformanceGuards As Boolean
    
    ' Using main Application instance for all operations
    
    ' Global tracking for all opened workbooks (for error cleanup)
    Dim globalOpenedWorkbooks As Collection
    Set globalOpenedWorkbooks = New Collection
    
    ' Track formula placement errors
    Dim formulaErrors As Long
    formulaErrors = 0
    
    On Error GoTo ErrHandler
    
    booEnablePerformanceGuards = True
    
    ' UX / Performance guards
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    
    If booEnablePerformanceGuards = True Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If
    
    
    ' 1) Empty temp folder (contents only)
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolderExists fso, strGradebooksTempFolderURL
    DeleteFolderContents fso, strGradebooksTempFolderURL
    Log logLines, "Cleared temp folder: " & strGradebooksTempFolderURL
    
    ' 2) Copy all contents from source -> temp
    EnsureFolderExists fso, strSourceFolderURL
    fso.CopyFolder strSourceFolderURL & "\*", strGradebooksTempFolderURL & "\"
    Log logLines, "Copied all contents from: " & strSourceFolderURL & " -> " & strGradebooksTempFolderURL
    
    ' 3) Resolve bimester folder
    EnsureFolderExists fso, strBimesterFolderURL
    Log logLines, "Bimester folder: " & strBimesterFolderURL
    
    ' 4) Count total .xlsx templates for progress tracking
    Dim templateCount As Long
    templateCount = 0
    Dim templatePath As String
    templatePath = Dir(strBimesterFolderURL & "\*.xlsx")
    
    ' First pass: count templates
    Do While Len(templatePath) > 0
        templateCount = templateCount + 1
        templatePath = Dir()
    Loop
    
    ' Reset for actual processing
    templatePath = Dir(strBimesterFolderURL & "\*.xlsx")
    
    ' Clear previous health report entries before starting new run
    ClearPreviousHealthReportEntries
    Log logLines, "Cleared previous health report entries for new run"
    
    ' Initialize ProgressBar
    Dim MyProgressbar As ProgressBar
    Set MyProgressbar = New ProgressBar
    
    With MyProgressbar
        .Title = "Generating Raw Gradebooks - " & strBimester
        .ExcelStatusBar = True
        .StartColour = rgbMediumSeaGreen
        .EndColour = rgbGreen
        .TotalActions = templateCount
    End With
    
    MyProgressbar.ShowBar
    Log logLines, "Starting processing of " & templateCount & " templates"
    
    ' Second pass: process templates
    Do While Len(templatePath) > 0
        Dim fullTemplatePath As String
        fullTemplatePath = strBimesterFolderURL & "\" & templatePath
        
        ' 4.1) Extract tag after "Grades-" and before "-Computers"
        Dim strGradeLevelTag As String
        strGradeLevelTag = GetBetween(templatePath, "Grades-", "-Computers")
        
        If Len(strGradeLevelTag) = 0 Then
            Log logLines, "SKIP (cannot parse grade tag): " & templatePath
            GoTo NextTemplate
        End If
        
        ' 4.2) Map to grade code
        Dim strGradeLevel As String
        strGradeLevel = MapGradeTagToCode(strGradeLevelTag)
        If Len(strGradeLevel) = 0 Then
            Log logLines, "SKIP (unknown grade mapping for tag '" & strGradeLevelTag & "'): " & templatePath
            GoTo NextTemplate
        End If
        
        ' 4.2.1) Filter by grade levels if provided
        If Not IsEmpty(gradeLevels) Then
            If Not IsGradeLevelIncluded(strGradeLevel, gradeLevels) Then
                Log logLines, "SKIP (grade level '" & strGradeLevel & "' not in filter list): " & templatePath
                GoTo NextTemplate
            End If
        End If
        
        Log logLines, "Processing template: " & templatePath & " | Tag='" & strGradeLevelTag & "' ? Code='" & strGradeLevel & "'"
        
        ' Update progress bar
        MyProgressbar.NextAction "Processing '" & templatePath & "'", True
        
        ' 4.3) Open the matching grade workbook(s) from each immediate subfolder
        Dim openedRefs As Collection
        Set openedRefs = New Collection  ' Collection of paths we opened (to close later)
        OpenMatchingFromSubfolders strBimesterFolderURL, strGradeLevel, openedRefs, globalOpenedWorkbooks, logLines
        
        ' 4.4) Open the template workbook in the main Application instance
        Dim wbTemplate As Object
        Set wbTemplate = GetOpenWorkbookByFullPath(fullTemplatePath)
        If wbTemplate Is Nothing Then
            ' Use the local path for opening, not any converted SharePoint URL
            Log logLines, "Opening template with local path: " & fullTemplatePath
            Set wbTemplate = Application.Workbooks.Open(fullTemplatePath)
            If Not wbTemplate Is Nothing Then
                Log logLines, "Template opened successfully. FullName: " & wbTemplate.FullName
                ' Track template workbook globally for error cleanup
                globalOpenedWorkbooks.Add fullTemplatePath
            Else
                Log logLines, "ERROR: Failed to open template: " & fullTemplatePath
            End If
        Else
            Log logLines, "Template already open in main instance: " & fullTemplatePath
        End If
        
        ' 4.4.1) Place formula in template if successfully opened
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            PlaceFormulaInTemplate wbTemplate, logLines
            If Err.Number <> 0 Then
                formulaErrors = formulaErrors + 1
                Log logLines, "ERROR placing formula in template: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.5) Replace formulas by values in the single sheet's rectangular range
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            ReplaceFormulasWithValues wbTemplate, logLines
            If Err.Number <> 0 Then
                Log logLines, "ERROR replacing formulas: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "SKIP replacing formulas - template workbook not available: " & templatePath
        End If
        
        ' 4.5.1) Clear zero grade values (cells that evaluate to 0)
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            ClearZeroGrades wbTemplate, logLines
            If Err.Number <> 0 Then
                Log logLines, "ERROR clearing zero grades: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "SKIP clearing zero grades - template workbook not available: " & templatePath
        End If
        
        ' 4.6) Open integrative activity workbook and place formula (after weekly grades)
        Dim integrativeActivityWb As Object
        Set integrativeActivityWb = Nothing
        
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            Set integrativeActivityWb = FindIntegrativeActivityWorkbook(wbTemplate, logLines)
            If Not integrativeActivityWb Is Nothing Then
                ' Track integrative activity workbook for cleanup
                globalOpenedWorkbooks.Add integrativeActivityWb.FullName
                Log logLines, "Added integrative activity workbook to global tracking (count=" & globalOpenedWorkbooks.Count & ")"
            End If
            On Error GoTo ErrHandler
        End If
        
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            PlaceIntegrativeActivityFormulaWithWorkbook wbTemplate, integrativeActivityWb, logLines
            If Err.Number <> 0 Then
                Log logLines, "ERROR placing integrative activity formula in template: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.6.1) Convert integrative activity formulas to values
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            ReplaceIntegrativeActivityFormulasWithValues wbTemplate, logLines
            If Err.Number <> 0 Then
                Log logLines, "ERROR replacing integrative activity formulas: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "SKIP replacing integrative activity formulas - template workbook not available: " & templatePath
        End If
        
        ' 4.7) Run health check on generated gradebook (optional)
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            RunHealthCheckOnGeneratedGradebook wbTemplate, templatePath
            If Err.Number <> 0 Then
                Log logLines, "WARN: Health check failed for " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.8) Save & close template using cloud-aware pattern
        If Not wbTemplate Is Nothing Then
            Log logLines, "Attempting to close template: " & wbTemplate.Name
            Log logLines, "Template FullName: " & wbTemplate.FullName
            Log logLines, "Template Saved property: " & wbTemplate.Saved
            Log logLines, "Template ReadOnly property: " & wbTemplate.ReadOnly
            
            CloseWorkbookSmart wbTemplate, logLines
            ' Remove template from global collection since it's now closed
            On Error Resume Next
            RemoveFromGlobalCollection globalOpenedWorkbooks, fullTemplatePath
            If Err.Number <> 0 Then
                Log logLines, "ERROR removing template from global collection: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "SKIP save/close - template workbook not available: " & templatePath
        End If
        
        ' 4.8.1) Close the integrative activity workbook
        If Not integrativeActivityWb Is Nothing Then
            On Error Resume Next
            integrativeActivityWb.Close SaveChanges:=False
            If Err.Number <> 0 Then
                Log logLines, "ERROR closing integrative activity workbook: " & integrativeActivityWb.Name & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            ' Remove from global collection since it's now closed
            On Error Resume Next
            RemoveFromGlobalCollection globalOpenedWorkbooks, integrativeActivityWb.FullName
            If Err.Number <> 0 Then
                Log logLines, "ERROR removing integrative activity workbook from global collection: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            Set integrativeActivityWb = Nothing
        End If
        
        ' 4.9) Close only the subfolder files that were opened in step 4.3
        CloseOpenedWorkbooks openedRefs, globalOpenedWorkbooks, logLines
        
NextTemplate:
        templatePath = Dir() ' next .xlsx
    Loop
    
    ' Process completed successfully
    Log logLines, "SUCCESS: Process completed - all workbooks closed properly"
    
    ' No cleanup needed - using main Application instance
    
    ' Complete progress bar
    On Error Resume Next
    MyProgressbar.Complete
    If Err.Number <> 0 Then
        Log logLines, "WARN: Error completing progress bar: " & Err.Description
        Err.Clear
    End If
    
    ' Close progress bar after completion
    On Error Resume Next
    MyProgressbar.Terminate
    If Err.Number <> 0 Then
        Log logLines, "WARN: Error terminating progress bar: " & Err.Description
        Err.Clear
    End If
    Set MyProgressbar = Nothing
    
    ' Flush log (while performance guards are still active)
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "GRB_Log"
    
    ' Wrap-up - restore performance guards AFTER logging
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    If prevCalc <> xlCalculationManual Then Application.Calculate
    If Err.Number <> 0 Then
        Log logLines, "WARN: Error during calculation restore: " & Err.Description
        Err.Clear
    End If
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    
    ' Process completed
    Dim finalMessage As String
    finalMessage = "GenerateRawGradebooks completed successfully."
    If formulaErrors > 0 Then
        finalMessage = finalMessage & " However, " & formulaErrors & " formula placement error(s) occurred. Check the log for details."
    Else
        finalMessage = finalMessage & " Check the log for details."
    End If
    
    On Error Resume Next
    MsgBox finalMessage, vbInformation, "Process Complete"
    If Err.Number <> 0 Then
        Log logLines, "WARN: Error showing completion message: " & Err.Description
        Err.Clear
    End If
    
    Exit Sub

ErrHandler:
    ' Best-effort restore
    On Error Resume Next
    
    ' Close progress bar if it exists
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Terminate
        Set MyProgressbar = Nothing
    End If
    
    ' Close all tracked workbooks before restoring settings
    CloseAllTrackedWorkbooks globalOpenedWorkbooks, logLines
    
    Log logLines, "FATAL: " & Err.Number & " - " & Err.Description
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "GRB_Log"
    
    ' Restore performance guards AFTER logging
    Application.Calculation = prevCalc
    If prevCalc <> xlCalculationManual Then Application.Calculate
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    
    MsgBox "GenerateRawGradebooks encountered an error: " & Err.Description, vbExclamation
End Sub

' ===========================
' Helpers
' ===========================

Private Sub ReplaceFormulasWithValues(ByVal wb As Object, ByRef logLines As Collection)
    ' Specs:
    ' - Single sheet in template
    ' - Rectangular range starts at C5
    ' - Last row: last non-empty cell in column B
    ' - Last column: prefer "Week ..." headers in row 3; fall back to dark/black fill
    ' - Replace only formulas in that area with their current values
    
    If wb Is Nothing Then
        Log logLines, "ERROR: ReplaceFormulasWithValues called with Nothing workbook object"
        Exit Sub
    End If
    
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping: " & wb.Name
        Exit Sub
    End If
    
    ' >>> Key change: use the same rule as PlaceFormulaInTemplate <<<
    lastCol = GetLastWeekColumnInRow(ws, 3)       ' try by "Week ..." header text
    If lastCol < 3 Then
        ' Fallback for older templates where headers are only styled (dark fill)
        lastCol = GetLastBlackBackgroundColInRow(ws, 3)
    End If
    If lastCol < 3 Then
        Log logLines, "INFO: No header columns detected in row 3 (text/fill). Skipping: " & wb.Name
        Exit Sub
    End If
    
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    
    ' Temporarily enable calculation to ensure formulas are calculated before replacement
    Dim tempCalc As XlCalculation
    tempCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate  ' Ensure all formulas are calculated
    
    ' Replace formulas with values
    Dim cell As Object  ' Late-bound Range
    Dim cnt As Long
    For Each cell In rng.Cells
        If cell.HasFormula Then
            cell.value = cell.value
            cnt = cnt + 1
        End If
    Next cell
    
    ' Restore original calculation mode
    Application.Calculation = tempCalc
    
    Log logLines, "Replaced " & cnt & " formulas with values in " & wb.Name & " | Range=" & rng.Address(External:=False)
End Sub

Private Sub ClearZeroGrades(ByVal wb As Object, ByRef logLines As Collection)
    ' Clears all cells with grade value that evaluates to 0 in the rectangular range
    ' Uses same range detection as ReplaceFormulasWithValues
    ' Also clears zero values in the integrative activity column
    
    If wb Is Nothing Then
        Log logLines, "ERROR: ClearZeroGrades called with Nothing workbook object"
        Exit Sub
    End If
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping zero grade clearing: " & wb.Name
        Exit Sub
    End If
    
    ' Use the same rule as ReplaceFormulasWithValues
    lastCol = GetLastWeekColumnInRow(ws, 3)       ' try by "Week ..." header text
    If lastCol < 3 Then
        ' Fallback for older templates where headers are only styled (dark fill)
        lastCol = GetLastBlackBackgroundColInRow(ws, 3)
    End If
    If lastCol < 3 Then
        Log logLines, "INFO: No header columns detected in row 3 (text/fill). Skipping zero grade clearing: " & wb.Name
        Exit Sub
    End If
    
    ' Clear zero grades in weekly columns
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    
    ' Clear cells that evaluate to 0
    Dim cell As Object  ' Late-bound Range
    Dim clearedCount As Long
    Dim cellValue As Variant
    
    For Each cell In rng.Cells
        cellValue = cell.Value
        ' Check if cell value evaluates to 0 (handles 0, 0.0, "0", etc.)
        If IsNumeric(cellValue) And CDbl(cellValue) = 0 Then
            cell.ClearContents
            clearedCount = clearedCount + 1
        End If
    Next cell
    
    Log logLines, "Cleared " & clearedCount & " zero grade cells in weekly columns for " & wb.Name & " | Range=" & rng.Address(External:=False)
    
    ' Also clear zero grades in integrative activity column
    Dim integrativeActivityCol As Long
    integrativeActivityCol = GetIntegrativeActivityColumn(ws)
    If integrativeActivityCol > 0 Then
        Dim iaRng As Object  ' Late-bound Range
        Set iaRng = ws.Range(ws.Cells(5, integrativeActivityCol), ws.Cells(lastRow, integrativeActivityCol))
        
        Dim iaClearedCount As Long
        For Each cell In iaRng.Cells
            cellValue = cell.Value
            ' Check if cell value evaluates to 0 (handles 0, 0.0, "0", etc.)
            If IsNumeric(cellValue) And CDbl(cellValue) = 0 Then
                cell.ClearContents
                iaClearedCount = iaClearedCount + 1
            End If
        Next cell
        
        Log logLines, "Cleared " & iaClearedCount & " zero grade cells in integrative activity column for " & wb.Name & " | Range=" & iaRng.Address(External:=False)
    Else
        Log logLines, "No integrative activity column found in " & wb.Name & " - skipping zero grade clearing for integrative activity"
    End If
End Sub

Private Function GetLastNonEmptyRowInColumn(ByVal ws As Object, ByVal colNum As Long) As Long
    Dim lastCell As Object  ' Late-bound Range
    Set lastCell = ws.Cells(ws.Rows.Count, colNum).End(xlUp)
    If Len(lastCell.value) = 0 And lastCell.row = 1 Then
        GetLastNonEmptyRowInColumn = 0
    Else
        GetLastNonEmptyRowInColumn = lastCell.row
    End If
End Function

Private Function IsBlackFill(ByVal cell As Object) As Boolean
    With cell.Interior
        If .pattern = xlNone Then Exit Function
        If .Color = vbBlack Or .ColorIndex = 1 Then
            IsBlackFill = True
        ElseIf .TintAndShade <> 0 Then
            ' treat very dark fills as black-ish
            Dim c As Long: c = .Color
            Dim r As Long: r = c Mod 256
            Dim g As Long: g = (c \ 256) Mod 256
            Dim b As Long: b = (c \ 65536) Mod 256
            IsBlackFill = (r + g + b) < 30
        End If
    End With
End Function

Private Function GetLastBlackBackgroundColInRow(ByVal ws As Object, ByVal rowNum As Long) As Long
    Dim lastUsedCol As Long, c As Long
    lastUsedCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    For c = lastUsedCol To 1 Step -1
        If IsBlackFill(ws.Cells(rowNum, c)) Then
            GetLastBlackBackgroundColInRow = c
            Exit Function
        End If
    Next c
    GetLastBlackBackgroundColInRow = 0
End Function

Private Function GetLastWeekColumnInRow(ByVal ws As Object, ByVal rowNum As Long) As Long
    Dim lastUsedCol As Long, c As Long
    lastUsedCol = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    For c = lastUsedCol To 1 Step -1
        Dim cellValue As String
        cellValue = CStr(ws.Cells(rowNum, c).Value)
        If Left(Trim(cellValue), 4) = "Week" Then
            GetLastWeekColumnInRow = c
            Exit Function
        End If
    Next c
    GetLastWeekColumnInRow = 0
End Function

Private Function GetIntegrativeActivityColumn(ByVal ws As Object) As Long
    ' Find the column containing "Actividad Integradora" in row 2
    Dim lastUsedCol As Long, c As Long
    lastUsedCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastUsedCol
        Dim cellValue As String
        cellValue = CStr(ws.Cells(2, c).Value)
        If InStr(1, cellValue, "Actividad Integradora", vbTextCompare) > 0 Then
            GetIntegrativeActivityColumn = c
            Exit Function
        End If
    Next c
    GetIntegrativeActivityColumn = 0
End Function

Private Function GetTotalColumnInIntegrativeActivity(ByVal wb As Object) As Long
    ' Find the column containing "Total" in row 5 of the "Actividad Integradora" worksheet
    If wb Is Nothing Then
        GetTotalColumnInIntegrativeActivity = 0
        Exit Function
    End If
    
    Dim ws As Object
    On Error Resume Next
    Set ws = wb.Worksheets("Actividad Integradora")
    If Err.Number <> 0 Or ws Is Nothing Then
        GetTotalColumnInIntegrativeActivity = 0
        Exit Function
    End If
    On Error GoTo 0
    
    Dim lastUsedCol As Long, c As Long
    lastUsedCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastUsedCol
        Dim cellValue As String
        cellValue = CStr(ws.Cells(5, c).Value)
        If InStr(1, cellValue, "Total", vbTextCompare) > 0 Then
            GetTotalColumnInIntegrativeActivity = c
            Exit Function
        End If
    Next c
    GetTotalColumnInIntegrativeActivity = 0
End Function

Private Sub PlaceFormulaInTemplate(ByVal wb As Object, ByRef logLines As Collection)
    ' Places the grade lookup formula in the template and copies it to the appropriate range
    
    If wb Is Nothing Then
        Log logLines, "ERROR: PlaceFormulaInTemplate called with Nothing workbook object"
        Exit Sub
    End If
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    ' Get the range dimensions
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping formula placement: " & wb.Name
        Exit Sub
    End If
    
    lastCol = GetLastWeekColumnInRow(ws, 3)  ' row 3 = 3
    If lastCol < 3 Then
        Log logLines, "INFO: No Week columns found in row 3. Skipping formula placement: " & wb.Name
        Exit Sub
    End If
    
    ' Define the formula by building it in parts to avoid line continuation limit
    Dim formula As String
    formula = "=LET("
    formula = formula & "name,$B5,"
    formula = formula & "week_label,C$3,"
    formula = formula & "name_to_grade_level,LAMBDA(grade_string,SWITCH(grade_string,"
    formula = formula & """First Grade_A"",""1A"","
    formula = formula & """First Grade_B"",""1B"","
    formula = formula & """Second Grade_A"",""2A"","
    formula = formula & """Second Grade_B"",""2B"","
    formula = formula & """Third Grade_A"",""3A"","
    formula = formula & """Third Grade_B"",""3B"","
    formula = formula & """Fourth Grade_A"",""4A"","
    formula = formula & """Fourth Grade_B"",""4B"","
    formula = formula & """Fifth Grade_A"",""5A"","
    formula = formula & """Fifth Grade_B"",""5B"","
    formula = formula & """Sixth Grade_A"",""6A"","
    formula = formula & """Sixth Grade_B"",""6B"","
    formula = formula & """Seventh Grade_A"",""7A"","
    formula = formula & """Eighth Grade_A"",""8A"","
    formula = formula & """Ninth Grade_A"",""9A"","
    formula = formula & """Tenth Grade_A"",""10A"","
    formula = formula & """Eleventh Grade_A"",""11A"","
    formula = formula & """Twelfth Grade_A"",""12A"","
    formula = formula & """Ciclo Tres Development Center_A"",""DC3A""," 
    formula = formula & """Ciclo Tres Development Center_B"",""DC3B""," 
    formula = formula & """Ciclo Tres Development Center_B secundaria"",""DC3B-SEC"")),"
    formula = formula & "grade_level_phrase,TEXTBEFORE(TEXTAFTER(CELL(""filename""),""Grades-""),""-Computers""),"
    formula = formula & "grade_level,name_to_grade_level(grade_level_phrase),"
    formula = formula & "week_number,TEXTAFTER(week_label,"" ""),"
    formula = formula & "fixed_week_number,IF(LEN(week_number)=2,"""",""0"")&week_number,"
    formula = formula & "parsed_name,TEXTAFTER(name,""- ""),"
    formula = formula & "clean_name,TRIM(SUBSTITUTE(parsed_name,"" ,"","","")),"
    formula = formula & "source_ws_xlsx,""'[Weekly Grade - W""&fixed_week_number&"" - ""&grade_level&"" - 2526.xlsx]Nota Semanal'"","
    formula = formula & "source_ws_xlsm,""'[Weekly Grade - W""&fixed_week_number&"" - ""&grade_level&"" - 2526.xlsm]Nota Semanal'"","
    formula = formula & "name_rng_xlsx,INDIRECT(source_ws_xlsx&""!$A:$A""),"
    formula = formula & "grade_rng_xlsx,INDIRECT(source_ws_xlsx&""!$H:$H""),"
    formula = formula & "name_rng_xlsm,INDIRECT(source_ws_xlsm&""!$A:$A""),"
    formula = formula & "grade_rng_xlsm,INDIRECT(source_ws_xlsm&""!$H:$H""),"
    formula = formula & "grade,IFERROR(XLOOKUP(clean_name,name_rng_xlsx,grade_rng_xlsx,0),XLOOKUP(clean_name,name_rng_xlsm,grade_rng_xlsm,0)),"
    formula = formula & "IFERROR(grade,0))"
    
    ' Place formula in C5
    On Error Resume Next
    ws.Range(FORMULA_START_CELL).Formula = formula
    If Err.Number <> 0 Then
        Log logLines, "ERROR: Failed to place formula in " & wb.Name & " | " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Copy formula to the rectangular range
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    
    On Error Resume Next
    ws.Range(FORMULA_START_CELL).Copy
    rng.PasteSpecial xlPasteFormulas
    If Err.Number <> 0 Then
        Log logLines, "ERROR copying formula to range in " & wb.Name & ": " & Err.Description
        Err.Clear
    Else
        Log logLines, "Formula placed on range " & rng.Address(External:=False) & " in " & wb.Name
    End If
    On Error GoTo 0
    
    ' Clear clipboard using the main Application instance
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Sub PlaceIntegrativeActivityFormula(ByVal wb As Object, ByRef logLines As Collection)
    ' Places the integrative activity grade lookup formula in the template
    ' This version finds and opens the integrative activity workbook automatically
    
    If wb Is Nothing Then
        Log logLines, "ERROR: PlaceIntegrativeActivityFormula called with Nothing workbook object"
        Exit Sub
    End If
    
    ' Find the integrative activity workbook
    Dim integrativeActivityWb As Object
    Set integrativeActivityWb = FindIntegrativeActivityWorkbook(wb, logLines)
    If integrativeActivityWb Is Nothing Then
        Log logLines, "INFO: No matching integrative activity workbook found. Skipping integrative activity formula placement: " & wb.Name
        Exit Sub
    End If
    
    ' Call the main function with the workbook
    PlaceIntegrativeActivityFormulaWithWorkbook wb, integrativeActivityWb, logLines
End Sub

Private Sub PlaceIntegrativeActivityFormulaWithWorkbook(ByVal wb As Object, ByVal integrativeActivityWb As Object, ByRef logLines As Collection)
    ' Places the integrative activity grade lookup formula in the template
    ' This version accepts the integrative activity workbook as a parameter
    
    If wb Is Nothing Then
        Log logLines, "ERROR: PlaceIntegrativeActivityFormulaWithWorkbook called with Nothing workbook object"
        Exit Sub
    End If
    
    If integrativeActivityWb Is Nothing Then
        Log logLines, "INFO: No integrative activity workbook provided. Skipping integrative activity formula placement: " & wb.Name
        Exit Sub
    End If
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    ' Find the integrative activity column
    Dim integrativeActivityCol As Long
    integrativeActivityCol = GetIntegrativeActivityColumn(ws)
    If integrativeActivityCol = 0 Then
        Log logLines, "INFO: No 'Actividad Integradora' column found in row 2. Skipping integrative activity formula placement: " & wb.Name
        Exit Sub
    End If
    
    ' Get the range dimensions
    Dim lastRow As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping integrative activity formula placement: " & wb.Name
        Exit Sub
    End If
    
    ' Find the Total column in the integrative activity workbook
    Dim totalCol As Long
    totalCol = GetTotalColumnInIntegrativeActivity(integrativeActivityWb)
    If totalCol = 0 Then
        Log logLines, "INFO: No 'Total' column found in integrative activity workbook. Skipping integrative activity formula placement: " & wb.Name
        Exit Sub
    End If
    
    ' Build the formula
    Dim formula As String
    Dim columnLetter As String
    columnLetter = GetColumnLetter(totalCol)
    
    formula = "=LET("
    formula = formula & "name,$B5,"
    formula = formula & "parsed_name,TEXTAFTER(name,""- ""),"
    formula = formula & "clean_name,TRIM(SUBSTITUTE(parsed_name,"" ,"","","")),"
    formula = formula & "source_ws,""'[" & integrativeActivityWb.Name & "]Actividad Integradora'"","
    formula = formula & "name_rng,INDIRECT(source_ws&""!$A:$A""),"
    formula = formula & "grade_rng,INDIRECT(source_ws&""!$" & columnLetter & ":$" & columnLetter & """),"
    formula = formula & "grade,XLOOKUP(clean_name,name_rng,grade_rng,0),"
    formula = formula & "IFERROR(grade,0))"
    
    ' Place formula in the first data cell of the integrative activity column
    Dim startCell As String
    startCell = GetColumnLetter(integrativeActivityCol) & "5"
    
    On Error Resume Next
    ws.Range(startCell).Formula = formula
    If Err.Number <> 0 Then
        Log logLines, "ERROR: Failed to place integrative activity formula in " & wb.Name & " | " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Copy formula to the range
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, integrativeActivityCol), ws.Cells(lastRow, integrativeActivityCol))
    
    On Error Resume Next
    ws.Range(startCell).Copy
    rng.PasteSpecial xlPasteFormulas
    If Err.Number <> 0 Then
        Log logLines, "ERROR copying integrative activity formula to range in " & wb.Name & ": " & Err.Description
        Err.Clear
    Else
        Log logLines, "Integrative activity formula placed on range " & rng.Address(External:=False) & " in " & wb.Name
    End If
    On Error GoTo 0
    
    ' Clear clipboard
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Function FindIntegrativeActivityWorkbook(ByVal templateWb As Object, ByRef logLines As Collection) As Object
    ' Find the matching integrative activity workbook for this template
    If templateWb Is Nothing Then
        Set FindIntegrativeActivityWorkbook = Nothing
        Exit Function
    End If
    
    ' Extract grade level from template filename
    Dim templateName As String
    templateName = templateWb.Name
    
    Dim gradeLevelTag As String
    gradeLevelTag = GetBetween(templateName, "Grades-", "-Computers")
    If Len(gradeLevelTag) = 0 Then
        Set FindIntegrativeActivityWorkbook = Nothing
        Exit Function
    End If
    
    Dim gradeLevel As String
    gradeLevel = MapGradeTagToCode(gradeLevelTag)
    If Len(gradeLevel) = 0 Then
        Set FindIntegrativeActivityWorkbook = Nothing
        Exit Function
    End If
    
    ' Look for integrative activity workbook in the "Integrative Activity" subfolder
    Dim bimesterFolder As String
    ' Extract bimester folder path from template path
    Dim templatePath As String
    templatePath = templateWb.FullName
    
    ' If template was opened via SharePoint URL, convert back to local path for folder operations
    If Left(templatePath, 4) = "http" Then
        Dim localPath As String
        localPath = SharePointUrlToLocal(templatePath)
        If localPath <> "" Then
            templatePath = localPath
            Log logLines, "Converted SharePoint URL to local path for folder operations: " & localPath
        Else
            Log logLines, "ERROR: Could not convert SharePoint URL to local path: " & templatePath
            Set FindIntegrativeActivityWorkbook = Nothing
            Exit Function
        End If
    End If
    
    Dim pathParts As Variant
    pathParts = Split(templatePath, "\")
    Dim bimesterFolderName As String
    If UBound(pathParts) >= 1 Then
        bimesterFolderName = pathParts(UBound(pathParts) - 1)
    End If
    
    ' Construct the integrative activity folder path
    Dim integrativeActivityFolder As String
    Dim lastSlashPos As Long
    lastSlashPos = InStrRev(templatePath, "\")
    If lastSlashPos > 0 Then
        integrativeActivityFolder = Left(templatePath, lastSlashPos - 1) & "\Integrative Activity"
    Else
        Log logLines, "ERROR: Could not find backslash in template path: " & templatePath
        Set FindIntegrativeActivityWorkbook = Nothing
        Exit Function
    End If
    
    Log logLines, "Looking for Integrative Activity folder: " & integrativeActivityFolder
    
    ' Look for the matching file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(integrativeActivityFolder) Then
        Log logLines, "Integrative Activity folder not found: " & integrativeActivityFolder
        Log logLines, "Template path used: " & templatePath
        Log logLines, "Last slash position: " & lastSlashPos
        Set FindIntegrativeActivityWorkbook = Nothing
        Exit Function
    End If
    
    Dim pattern As String
    pattern = "Integrative Activity - " & gradeLevel & " -"
    
    Dim file As Object
    For Each file In fso.GetFolder(integrativeActivityFolder).Files
        If LCase$(Right$(file.Name, 5)) = ".xlsm" Or LCase$(Right$(file.Name, 5)) = ".xlsx" Then
            If InStr(1, file.Name, pattern, vbTextCompare) > 0 Then
                ' Check if this workbook is already open
                Dim wb As Object
                Set wb = GetOpenWorkbookByFullPath(file.path)
                If Not wb Is Nothing Then
                    Set FindIntegrativeActivityWorkbook = wb
                    Log logLines, "Found already open integrative activity workbook: " & file.Name
                    Exit Function
                Else
                    ' Open the workbook if not already open
                    On Error Resume Next
                    Set wb = Application.Workbooks.Open(file.path)
                    If Err.Number <> 0 Then
                        Log logLines, "ERROR: Failed to open integrative activity workbook: " & file.Name & " | " & Err.Description
                        Err.Clear
                        Set FindIntegrativeActivityWorkbook = Nothing
                        Exit Function
                    End If
                    On Error GoTo 0
                    
                    If Not wb Is Nothing Then
                        Set FindIntegrativeActivityWorkbook = wb
                        Log logLines, "Opened integrative activity workbook: " & file.Name
                        Exit Function
                    End If
                End If
            End If
        End If
    Next file
    
    Set FindIntegrativeActivityWorkbook = Nothing
End Function

Private Function GetColumnLetter(ByVal columnNumber As Long) As String
    ' Convert column number to Excel column letter (1=A, 2=B, etc.)
    Dim dividend As Long
    Dim columnName As String
    Dim modulo As Long
    
    dividend = columnNumber
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = Int((dividend - modulo) / 26)
    Loop
    
    GetColumnLetter = columnName
End Function

Private Sub CloseWorkbookSmart(ByVal wb As Workbook, ByRef logLines As Collection)
    ' Cloud-aware workbook closing pattern based on ChatGPT recommendations
    Dim isCloud As Boolean
    Dim autoSaveOn As Boolean
    
    On Error GoTo CleanFail
    
    ' Detect SharePoint/OneDrive URL (opened from cloud, not local sync folder)
    isCloud = (LCase$(Left$(wb.FullName, 8)) = "https://") Or (InStr(1, wb.FullName, "sharepoint.com", vbTextCompare) > 0)
    
    ' AutoSave exists in Microsoft 365
    On Error Resume Next
    autoSaveOn = wb.AutoSaveOn
    On Error GoTo CleanFail
    
    Log logLines, "Cloud detection: " & isCloud & " | AutoSave: " & autoSaveOn
    
    ' Never saved -> you must SaveAs somewhere
    If wb.Path = "" Then
        Log logLines, "Workbook never saved, saving to temp location"
        wb.SaveAs ThisWorkbook.Path & "\Unsaved_" & wb.Name
        wb.Close SaveChanges:=False
        Log logLines, "Workbook saved and closed from temp location"
        Exit Sub
    End If
    
    If isCloud Then
        ' Best practice for cloud files: avoid explicit Save; let Close do the commit
        Log logLines, "Closing cloud workbook with SaveChanges:=True"
        wb.Close SaveChanges:=True
        Log logLines, "Cloud workbook closed successfully"
    Else
        ' Local path (including OneDrive *synced* local file): explicit Save is fine
        Log logLines, "Saving local workbook explicitly"
        wb.Save
        Log logLines, "Local workbook saved, closing with SaveChanges:=False"
        wb.Close SaveChanges:=False
        Log logLines, "Local workbook saved and closed successfully"
    End If
    Exit Sub

CleanFail:
    ' If Close failed due to a transient cloud error, make it non-fatal
    Log logLines, "ERROR in CloseWorkbookSmart: " & Err.Number & " - " & Err.Description
    On Error Resume Next
    wb.Close SaveChanges:=False
    Log logLines, "Emergency close completed (SaveChanges:=False)"
End Sub

Private Sub OpenMatchingFromSubfolders(ByVal bimesterFolder As String, ByVal gradeCode As String, _
                                       ByRef openedRefs As Collection, ByRef globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    Dim fso As Object, folder As Object, subf As Object, fil As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(bimesterFolder)

    Dim pattern As String
    pattern = "- " & gradeCode & " -"

    For Each subf In folder.SubFolders
        Dim openedAny As Boolean
        ' Enumerate files with FSO to avoid nested Dir()
        For Each fil In subf.Files
            ' Only .xlsx
            If LCase$(Right$(fil.Name, 5)) = ".xlsx" Or LCase$(Right$(fil.Name, 5)) = ".xlsm" Then
                If InStr(1, fil.Name, pattern, vbTextCompare) > 0 Then
                    Dim fullPath As String
                    fullPath = fil.path

                    ' Open only if not already open in the main instance; track only ones we open now
                    If GetOpenWorkbookByFullPath(fullPath) Is Nothing Then
                        Dim wb As Object
                        Set wb = Application.Workbooks.Open(fullPath)
                        If Not wb Is Nothing Then
                            openedRefs.Add fullPath
                            globalOpenedWorkbooks.Add fullPath  ' Also track globally for error cleanup
                            openedAny = True
                            Log logLines, "Added to openedRefs (count=" & openedRefs.Count & ") and globalOpenedWorkbooks (count=" & globalOpenedWorkbooks.Count & ")"
                        End If
                    Else
                        ' Already open in the main instance: do NOT close later
                        Log logLines, "Already open in main instance: " & fullPath
                    End If
                End If
            End If
        Next fil

        If Not openedAny Then
            Log logLines, "No matching file in subfolder: " & subf.path & " (pattern '" & pattern & "')"
        End If
    Next subf
End Sub

Private Sub CloseOpenedWorkbooks(ByVal openedRefs As Collection, ByRef globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    Dim i As Long
    For i = openedRefs.Count To 1 Step -1
        Dim p As String
        p = CStr(openedRefs(i))
        Dim wb As Object
        Set wb = GetOpenWorkbookByFullPath(p)
        If Not wb Is Nothing Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            If Err.Number <> 0 Then
                Log logLines, "ERROR closing data file: " & p & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
            ' Remove from global collection since it's now closed
            RemoveFromGlobalCollection globalOpenedWorkbooks, p
        End If
        openedRefs.Remove i
    Next i
End Sub

Private Sub CloseAllTrackedWorkbooks(ByVal globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    ' Close all workbooks that were opened during the process (for error cleanup)
    Dim i As Long
    For i = globalOpenedWorkbooks.Count To 1 Step -1
        Dim p As String
        p = CStr(globalOpenedWorkbooks(i))
        Dim wb As Object
        Set wb = GetOpenWorkbookByFullPath(p)
        If Not wb Is Nothing Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            If Err.Number <> 0 Then
                Log logLines, "ERROR CLEANUP - Failed to close: " & p & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
        globalOpenedWorkbooks.Remove i
    Next i
End Sub

Private Sub RemoveFromGlobalCollection(ByRef globalOpenedWorkbooks As Collection, ByVal pathToRemove As String)
    ' Remove a specific path from the global collection
    Dim i As Long
    For i = globalOpenedWorkbooks.Count To 1 Step -1
        If StrComp(CStr(globalOpenedWorkbooks(i)), pathToRemove, vbTextCompare) = 0 Then
            globalOpenedWorkbooks.Remove i
            Exit For
        End If
    Next i
End Sub


Private Function GetOpenWorkbookByFullPath(ByVal targetPath As String) As Workbook
    Dim wb As Workbook
    Dim localPath As String
    Dim sharePointPath As String
    
    ' Try direct match first
    For Each wb In Application.Workbooks
        On Error Resume Next
        ' Some workbooks may not expose FullName safely; ignore errors
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullPath = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
    
    ' If no direct match, try to convert local OneDrive path to SharePoint URL
    If InStr(targetPath, "OneDrive - ABC BILINGUAL SCHOOL") > 0 Then
        localPath = targetPath
        sharePointPath = ConvertLocalPathToSharePointURL(localPath)
        
        For Each wb In Application.Workbooks
            On Error Resume Next
            If StrComp(wb.FullName, sharePointPath, vbTextCompare) = 0 Then
                Set GetOpenWorkbookByFullPath = wb
                Exit Function
            End If
            On Error GoTo 0
        Next wb
    End If
    
    ' If still no match, try to convert SharePoint URL to local path
    For Each wb In Application.Workbooks
        On Error Resume Next
        If InStr(wb.FullName, "sharepoint.com") > 0 Then
            localPath = ConvertSharePointURLToLocalPath(wb.FullName)
            If StrComp(localPath, targetPath, vbTextCompare) = 0 Then
                Set GetOpenWorkbookByFullPath = wb
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next wb
    
    Set GetOpenWorkbookByFullPath = Nothing
End Function


' Helper function to convert local OneDrive path to SharePoint URL
Private Function ConvertLocalPathToSharePointURL(ByVal localPath As String) As String
    Dim sharePointPath As String
    Dim relativePath As String
    
    ' Extract the relative path after "OneDrive - ABC BILINGUAL SCHOOL"
    relativePath = Mid(localPath, InStr(localPath, "OneDrive - ABC BILINGUAL SCHOOL") + Len("OneDrive - ABC BILINGUAL SCHOOL") + 1)
    
    ' Convert backslashes to forward slashes
    relativePath = Replace(relativePath, "\", "/")
    
    ' Construct SharePoint URL
    sharePointPath = "https://abcbilingualschool-my.sharepoint.com/personal/jorge_lopez_abcbilingualschool_edu_sv/Documents/" & relativePath
    
    ConvertLocalPathToSharePointURL = sharePointPath
End Function

' Helper function to convert SharePoint URL to local OneDrive path
Private Function ConvertSharePointURLToLocalPath(ByVal sharePointURL As String) As String
    Dim localPath As String
    Dim relativePath As String
    
    ' Extract the relative path after "/Documents/"
    relativePath = Mid(sharePointURL, InStr(sharePointURL, "/Documents/") + Len("/Documents/"))
    
    ' Convert forward slashes to backslashes
    relativePath = Replace(relativePath, "/", "\")
    
    ' Construct local OneDrive path
    localPath = "C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\" & relativePath
    
    ConvertSharePointURLToLocalPath = localPath
End Function

' Helper function to convert SharePoint URL to local path (for Files On-Demand compatibility)
Private Function SharePointUrlToLocal(ByVal spUrl As String) As String
    Dim root As String, rel As String, oneDriveRoot As String, p As Long
    
    ' Try OneDrive business root (e.g., "C:\Users\<you>\OneDrive - YourOrg")
    oneDriveRoot = Environ("OneDriveCommercial")
    If oneDriveRoot = "" Then oneDriveRoot = Environ("OneDrive")
    
    ' Look for the root library in SharePoint URL
    p = InStr(1, spUrl, "/Documents/", vbTextCompare)
    If p > 0 And oneDriveRoot <> "" Then
        rel = Mid(spUrl, p + Len("/Documents/"))
        rel = Replace(rel, "%20", " ")
        ' Convert ALL forward slashes to backslashes
        rel = Replace(rel, "/", "\")
        SharePointUrlToLocal = oneDriveRoot & "\" & rel
    Else
        SharePointUrlToLocal = ""  ' unknown mapping; handle accordingly
    End If
End Function

Private Function GetBetween(ByVal text As String, ByVal after As String, ByVal before As String) As String
    Dim p1 As Long, p2 As Long, startPos As Long
    p1 = InStr(1, text, after, vbTextCompare)
    If p1 = 0 Then Exit Function
    startPos = p1 + Len(after)
    p2 = InStr(startPos, text, before, vbTextCompare)
    If p2 = 0 Then Exit Function
    GetBetween = Mid$(text, startPos, p2 - startPos)
End Function

Private Function IsGradeLevelIncluded(ByVal gradeLevel As String, ByVal gradeLevels As Variant) As Boolean
    ' Check if the given grade level is included in the provided list
    ' gradeLevels can be a single string or an array of strings
    
    If IsEmpty(gradeLevels) Then
        IsGradeLevelIncluded = True
        Exit Function
    End If
    
    Dim i As Long
    If IsArray(gradeLevels) Then
        ' Handle array of grade levels
        For i = LBound(gradeLevels) To UBound(gradeLevels)
            If StrComp(CStr(gradeLevels(i)), gradeLevel, vbTextCompare) = 0 Then
                IsGradeLevelIncluded = True
                Exit Function
            End If
        Next i
    Else
        ' Handle single grade level
        If StrComp(CStr(gradeLevels), gradeLevel, vbTextCompare) = 0 Then
            IsGradeLevelIncluded = True
            Exit Function
        End If
    End If
    
    IsGradeLevelIncluded = False
End Function

Private Function MapGradeTagToCode(ByVal tag As String) As String
    ' Handles:
    '  - "First Grade_A" ? "1A", "First Grade_B" ? "1B", ..., "Twelfth Grade_A/B" ? "12A/12B"
    '  - "Ciclo Tres Development Center_A" ? "DC3A"
    '  - "Ciclo Tres Development Center_B" ? "DC3B"
    '  - "Ciclo Tres Development Center_B secundaria" ? "DC3B-SEC"
    
    Dim sec As String
    Dim suffix As String
    
    ' Special DC3 cases first
    If InStr(1, tag, "Ciclo Tres Development Center", vbTextCompare) = 1 Then
        If Right$(tag, 12) = "_B secundaria" Then
            MapGradeTagToCode = "DC3B-SEC"
            Exit Function
        End If
        
        sec = Right$(tag, 2) ' expects "_A" or "_B"
        If sec = "_A" Or sec = "_B" Then
            suffix = Right$(sec, 1) ' "A" or "B"
            MapGradeTagToCode = "DC3" & suffix
            Exit Function
        End If
        
        MapGradeTagToCode = ""
        Exit Function
    End If
    
    ' Regular grade levels
    sec = Right$(tag, 2) ' expects "_A" or "_B"
    If Not (sec = "_A" Or sec = "_B") Then
        MapGradeTagToCode = ""
        Exit Function
    End If
    
    suffix = Right$(sec, 1) ' "A" or "B"
    
    Dim gradeWord As String
    gradeWord = Split(tag, " ")(0) ' "First", "Second", ...
    
    Dim n As Long
    n = GradeWordToNumber(gradeWord)
    If n = 0 Then
        MapGradeTagToCode = ""
        Exit Function
    End If
    
    MapGradeTagToCode = CStr(n) & suffix
End Function

Private Function GradeWordToNumber(ByVal word As String) As Long
    Select Case LCase$(word)
        Case "first":   GradeWordToNumber = 1
        Case "second":  GradeWordToNumber = 2
        Case "third":   GradeWordToNumber = 3
        Case "fourth":  GradeWordToNumber = 4
        Case "fifth":   GradeWordToNumber = 5
        Case "sixth":   GradeWordToNumber = 6
        Case "seventh": GradeWordToNumber = 7
        Case "eighth":  GradeWordToNumber = 8
        Case "ninth":   GradeWordToNumber = 9
        Case "tenth":   GradeWordToNumber = 10
        Case "eleventh": GradeWordToNumber = 11
        Case "twelfth": GradeWordToNumber = 12
        Case Else:      GradeWordToNumber = 0
    End Select
End Function

Private Sub EnsureFolderExists(ByVal fso As Object, ByVal path As String)
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub

Private Sub DeleteFolderContents(ByVal fso As Object, ByVal folderPath As String)
    Dim f As Object
    Dim fld As Object
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    With fso.GetFolder(folderPath)
        For Each f In .Files
            On Error Resume Next
            f.Delete True
            On Error GoTo 0
        Next f
        For Each fld In .SubFolders
            On Error Resume Next
            fso.DeleteFolder fld.path, True
            On Error GoTo 0
        Next fld
    End With
End Sub

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




' ===========================
' Logging
' ===========================
Private Sub Log(ByRef logLines As Collection, ByVal msg As String)
    Dim timestampedMsg As String
    timestampedMsg = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    
    ' Add to collection for worksheet logging
    logLines.Add timestampedMsg
    
    ' Write to Immediate window in real-time for user feedback
    Debug.Print timestampedMsg
End Sub

Private Sub DumpLogToImmediate(ByVal logLines As Collection)
    ' Since we now write to Immediate window in real-time, just add a separator
    Debug.Print String(60, "-")
    Debug.Print "GenerateRawGradebooks LOG COMPLETE @ " & Now
    Debug.Print String(60, "-")
End Sub

Private Sub DumpLogToSheet(ByVal logLines As Collection, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        If Err.Number <> 0 Then
            ' If we can't create the worksheet, exit the function
            Exit Sub
        End If
        ws.Name = sheetName
    End If
    
    ' Check if ws is still valid before using it
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ws.Range("A1").value = "Timestamp"
    ws.Range("B1").value = "Message"
    ws.Range("A1:B1").Font.Bold = True
    
    Dim i As Long
    For i = 1 To logLines.Count
        ws.Cells(i + 1, 1).value = Split(logLines(i), "  ")(0)
        ws.Cells(i + 1, 2).value = Mid$(logLines(i), Len(Split(logLines(i), "  ")(0)) + 3)
    Next i
    
    ws.Columns("A:B").EntireColumn.AutoFit
End Sub

' ===========================
' Test Functions for Refactored GenerateRawGradebooks
' ===========================

Public Sub TestGenerateRawGradebooksWithFiltering()
    ' Test function to demonstrate the new grade level filtering functionality
    ' This is for testing purposes only - remove or comment out in production
    
    Dim testGradeLevels As Variant
    
    ' Test 1: Process all grade levels (original behavior)
    Debug.Print "=== Test 1: Process all grade levels ==="
    ' GenerateRawGradebooks "B1"  ' Uncomment to test with actual data
    
    ' Test 2: Process only specific grade levels (single grade)
    Debug.Print "=== Test 2: Process only 1A ==="
    testGradeLevels = "1A"
    ' GenerateRawGradebooks "B1", testGradeLevels  ' Uncomment to test with actual data
    
    ' Test 3: Process multiple grade levels (array)
    Debug.Print "=== Test 3: Process 1A, 2B, and DC3A ==="
    testGradeLevels = Array("1A", "2B", "DC3A")
    ' GenerateRawGradebooks "B1", testGradeLevels  ' Uncomment to test with actual data
    
    ' Test 4: Process DC3 special grade levels
    Debug.Print "=== Test 4: Process DC3 grade levels ==="
    testGradeLevels = Array("DC3A", "DC3B", "DC3BSEC")
    ' GenerateRawGradebooks "B1", testGradeLevels  ' Uncomment to test with actual data
    
    Debug.Print "Test functions completed. Uncomment the actual function calls to test with real data."
End Sub

Public Sub TestIntegrativeActivityOnly(ByVal strBimester As String, Optional ByVal gradeLevels As Variant = Empty)
    ' Test function to process ONLY integrative activities, skipping weekly grade processing
    ' This is for testing purposes only - remove or comment out in production
    
    Dim strGradebooksTempFolderURL As String
    Dim strSourceFolderURL As String
    Dim strBimesterFolderURL As String
    
    ' ==== PLACEHOLDERS: set these three before running ====
    strGradebooksTempFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Temp_Grades\")
    strSourceFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Grades\")
    strBimesterFolderURL = JoinPath(strGradebooksTempFolderURL, strBimester)   ' e.g., B1

    ' ======================================================
    ' Initialize logging for health check validation
    ' ======================================================
    Dim logLines As Collection
    Set logLines = New Collection

    ' ======================================================
    ' Health Check Validation - SKIPPED FOR TESTING
    ' ======================================================
    Log logLines, "TEST MODE: Skipping health check validation for integrative activity testing"
    
    Dim fso As Object
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEvents As Boolean
    Dim booEnablePerformanceGuards As Boolean
    
    ' Global tracking for all opened workbooks (for error cleanup)
    Dim globalOpenedWorkbooks As Collection
    Set globalOpenedWorkbooks = New Collection
    
    ' Track formula placement errors
    Dim formulaErrors As Long
    formulaErrors = 0
    
    On Error GoTo ErrHandler
    
    booEnablePerformanceGuards = True
    
    ' UX / Performance guards
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEvents = Application.EnableEvents
    
    If booEnablePerformanceGuards = True Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    End If
    
    ' 1) Empty temp folder (contents only)
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolderExists fso, strGradebooksTempFolderURL
    DeleteFolderContents fso, strGradebooksTempFolderURL
    Log logLines, "TEST MODE: Cleared temp folder: " & strGradebooksTempFolderURL
    
    ' 2) Copy all contents from source -> temp
    EnsureFolderExists fso, strSourceFolderURL
    fso.CopyFolder strSourceFolderURL & "\*", strGradebooksTempFolderURL & "\"
    Log logLines, "TEST MODE: Copied all contents from: " & strSourceFolderURL & " -> " & strGradebooksTempFolderURL
    
    ' 3) Resolve bimester folder
    EnsureFolderExists fso, strBimesterFolderURL
    Log logLines, "TEST MODE: Bimester folder: " & strBimesterFolderURL
    
    ' 4) Count total .xlsx templates for progress tracking
    Dim templateCount As Long
    templateCount = 0
    Dim templatePath As String
    templatePath = Dir(strBimesterFolderURL & "\*.xlsx")
    
    ' First pass: count templates
    Do While Len(templatePath) > 0
        templateCount = templateCount + 1
        templatePath = Dir()
    Loop
    
    ' Reset for actual processing
    templatePath = Dir(strBimesterFolderURL & "\*.xlsx")
    
    ' Clear previous health report entries before starting new run
    ClearPreviousHealthReportEntries
    Log logLines, "TEST MODE: Cleared previous health report entries for new run"
    
    ' Initialize ProgressBar
    Dim MyProgressbar As ProgressBar
    Set MyProgressbar = New ProgressBar
    
    With MyProgressbar
        .Title = "TEST: Integrative Activity Only - " & strBimester
        .ExcelStatusBar = True
        .StartColour = rgbMediumSeaGreen
        .EndColour = rgbGreen
        .TotalActions = templateCount
    End With
    
    MyProgressbar.ShowBar
    Log logLines, "TEST MODE: Starting processing of " & templateCount & " templates (integrative activity only)"
    
    ' Second pass: process templates
    Do While Len(templatePath) > 0
        Dim fullTemplatePath As String
        fullTemplatePath = strBimesterFolderURL & "\" & templatePath
        
        ' 4.1) Extract tag after "Grades-" and before "-Computers"
        Dim strGradeLevelTag As String
        strGradeLevelTag = GetBetween(templatePath, "Grades-", "-Computers")
        
        If Len(strGradeLevelTag) = 0 Then
            Log logLines, "TEST MODE: SKIP (cannot parse grade tag): " & templatePath
            GoTo NextTemplate
        End If
        
        ' 4.2) Map to grade code
        Dim strGradeLevel As String
        strGradeLevel = MapGradeTagToCode(strGradeLevelTag)
        If Len(strGradeLevel) = 0 Then
            Log logLines, "TEST MODE: SKIP (unknown grade mapping for tag '" & strGradeLevelTag & "'): " & templatePath
            GoTo NextTemplate
        End If
        
        ' 4.2.1) Filter by grade levels if provided
        If Not IsEmpty(gradeLevels) Then
            If Not IsGradeLevelIncluded(strGradeLevel, gradeLevels) Then
                Log logLines, "TEST MODE: SKIP (grade level '" & strGradeLevel & "' not in filter list): " & templatePath
                GoTo NextTemplate
            End If
        End If
        
        Log logLines, "TEST MODE: Processing template: " & templatePath & " | Tag='" & strGradeLevelTag & "' ? Code='" & strGradeLevel & "'"
        
        ' Update progress bar
        MyProgressbar.NextAction "TEST: Processing '" & templatePath & "' (Integrative Activity Only)", True
        
        ' SKIP: Weekly grade processing (steps 4.3, 4.4, 4.5, 4.5.1)
        Log logLines, "TEST MODE: Skipping weekly grade processing for: " & templatePath
        
        ' 4.4) Open the template workbook in the main Application instance
        Dim wbTemplate As Object
        Set wbTemplate = GetOpenWorkbookByFullPath(fullTemplatePath)
        If wbTemplate Is Nothing Then
            ' Use the local path for opening, not any converted SharePoint URL
            Log logLines, "TEST MODE: Opening template with local path: " & fullTemplatePath
            Set wbTemplate = Application.Workbooks.Open(fullTemplatePath)
            If Not wbTemplate Is Nothing Then
                Log logLines, "TEST MODE: Template opened successfully. FullName: " & wbTemplate.FullName
                ' Track template workbook globally for error cleanup
                globalOpenedWorkbooks.Add fullTemplatePath
            Else
                Log logLines, "TEST MODE: ERROR: Failed to open template: " & fullTemplatePath
            End If
        Else
            Log logLines, "TEST MODE: Template already open in main instance: " & fullTemplatePath
        End If
        
        ' 4.6) Open integrative activity workbook and place formula (TEST MODE - ONLY THIS STEP)
        Dim integrativeActivityWb As Object
        Set integrativeActivityWb = Nothing
        
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            Set integrativeActivityWb = FindIntegrativeActivityWorkbook(wbTemplate, logLines)
            If Not integrativeActivityWb Is Nothing Then
                ' Track integrative activity workbook for cleanup
                globalOpenedWorkbooks.Add integrativeActivityWb.FullName
                Log logLines, "TEST MODE: Added integrative activity workbook to global tracking (count=" & globalOpenedWorkbooks.Count & ")"
            End If
            On Error GoTo ErrHandler
        End If
        
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            PlaceIntegrativeActivityFormulaWithWorkbook wbTemplate, integrativeActivityWb, logLines
            If Err.Number <> 0 Then
                formulaErrors = formulaErrors + 1
                Log logLines, "TEST MODE: ERROR placing integrative activity formula in template: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.7) Convert integrative activity formulas to values
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            ReplaceIntegrativeActivityFormulasWithValues wbTemplate, logLines
            If Err.Number <> 0 Then
                Log logLines, "TEST MODE: ERROR replacing integrative activity formulas: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "TEST MODE: SKIP replacing integrative activity formulas - template workbook not available: " & templatePath
        End If
        
        ' 4.8) Clear zero grades in integrative activity column only
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            ClearIntegrativeActivityZeroGrades wbTemplate, logLines
            If Err.Number <> 0 Then
                Log logLines, "TEST MODE: ERROR clearing integrative activity zero grades: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "TEST MODE: SKIP clearing integrative activity zero grades - template workbook not available: " & templatePath
        End If
        
        ' SKIP: Health check (step 4.7)
        Log logLines, "TEST MODE: Skipping health check for: " & templatePath
        
        ' 4.8) Save & close template using cloud-aware pattern
        If Not wbTemplate Is Nothing Then
            Log logLines, "TEST MODE: Attempting to close template: " & wbTemplate.Name
            Log logLines, "TEST MODE: Template FullName: " & wbTemplate.FullName
            Log logLines, "TEST MODE: Template Saved property: " & wbTemplate.Saved
            Log logLines, "TEST MODE: Template ReadOnly property: " & wbTemplate.ReadOnly
            
            CloseWorkbookSmart wbTemplate, logLines
            ' Remove template from global collection since it's now closed
            On Error Resume Next
            RemoveFromGlobalCollection globalOpenedWorkbooks, fullTemplatePath
            If Err.Number <> 0 Then
                Log logLines, "TEST MODE: ERROR removing template from global collection: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        Else
            Log logLines, "TEST MODE: SKIP save/close - template workbook not available: " & templatePath
        End If
        
        ' 4.8.1) Close the integrative activity workbook
        If Not integrativeActivityWb Is Nothing Then
            On Error Resume Next
            integrativeActivityWb.Close SaveChanges:=False
            If Err.Number <> 0 Then
                Log logLines, "TEST MODE: ERROR closing integrative activity workbook: " & integrativeActivityWb.Name & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            ' Remove from global collection since it's now closed
            On Error Resume Next
            RemoveFromGlobalCollection globalOpenedWorkbooks, integrativeActivityWb.FullName
            If Err.Number <> 0 Then
                Log logLines, "TEST MODE: ERROR removing integrative activity workbook from global collection: " & Err.Number & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
            Set integrativeActivityWb = Nothing
        End If
        
NextTemplate:
        templatePath = Dir() ' next .xlsx
    Loop
    
    ' Process completed successfully
    Log logLines, "TEST MODE: SUCCESS: Integrative activity processing completed - all workbooks closed properly"
    
    ' Complete progress bar
    On Error Resume Next
    MyProgressbar.Complete
    If Err.Number <> 0 Then
        Log logLines, "TEST MODE: WARN: Error completing progress bar: " & Err.Description
        Err.Clear
    End If
    
    ' Close progress bar after completion
    On Error Resume Next
    MyProgressbar.Terminate
    If Err.Number <> 0 Then
        Log logLines, "TEST MODE: WARN: Error terminating progress bar: " & Err.Description
        Err.Clear
    End If
    Set MyProgressbar = Nothing
    
    ' Flush log (while performance guards are still active)
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "IA_Test_Log"
    
    ' Wrap-up - restore performance guards AFTER logging
    On Error Resume Next
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    If prevCalc <> xlCalculationManual Then Application.Calculate
    If Err.Number <> 0 Then
        Log logLines, "TEST MODE: WARN: Error during calculation restore: " & Err.Description
        Err.Clear
    End If
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    
    ' Process completed
    Dim finalMessage As String
    finalMessage = "TEST MODE: Integrative Activity processing completed successfully."
    If formulaErrors > 0 Then
        finalMessage = finalMessage & " However, " & formulaErrors & " formula placement error(s) occurred. Check the log for details."
    Else
        finalMessage = finalMessage & " Check the IA_Test_Log sheet for details."
    End If
    
    On Error Resume Next
    MsgBox finalMessage, vbInformation, "TEST: Process Complete"
    If Err.Number <> 0 Then
        Log logLines, "TEST MODE: WARN: Error showing completion message: " & Err.Description
        Err.Clear
    End If
    
    Exit Sub

ErrHandler:
    ' Best-effort restore
    On Error Resume Next
    
    ' Close progress bar if it exists
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Terminate
        Set MyProgressbar = Nothing
    End If
    
    ' Close all tracked workbooks before restoring settings
    CloseAllTrackedWorkbooks globalOpenedWorkbooks, logLines
    
    Log logLines, "TEST MODE: FATAL: " & Err.Number & " - " & Err.Description
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "IA_Test_Log"
    
    ' Restore performance guards AFTER logging
    Application.Calculation = prevCalc
    If prevCalc <> xlCalculationManual Then Application.Calculate
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    
    MsgBox "TEST MODE: Integrative Activity processing encountered an error: " & Err.Description, vbExclamation
End Sub

Private Sub ReplaceIntegrativeActivityFormulasWithValues(ByVal wb As Object, ByRef logLines As Collection)
    ' Converts only the integrative activity formulas to values (for testing)
    
    If wb Is Nothing Then
        Log logLines, "ERROR: ReplaceIntegrativeActivityFormulasWithValues called with Nothing workbook object"
        Exit Sub
    End If
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Dim lastRow As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping integrative activity formula conversion: " & wb.Name
        Exit Sub
    End If
    
    ' Find the integrative activity column
    Dim integrativeActivityCol As Long
    integrativeActivityCol = GetIntegrativeActivityColumn(ws)
    If integrativeActivityCol = 0 Then
        Log logLines, "INFO: No 'Actividad Integradora' column found in row 2. Skipping integrative activity formula conversion: " & wb.Name
        Exit Sub
    End If
    
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, integrativeActivityCol), ws.Cells(lastRow, integrativeActivityCol))
    
    ' Temporarily enable calculation to ensure formulas are calculated before replacement
    Dim tempCalc As XlCalculation
    tempCalc = Application.Calculation
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate  ' Ensure all formulas are calculated
    
    ' Replace formulas with values
    Dim cell As Object  ' Late-bound Range
    Dim cnt As Long
    For Each cell In rng.Cells
        If cell.HasFormula Then
            cell.value = cell.value
            cnt = cnt + 1
        End If
    Next cell
    
    ' Restore original calculation mode
    Application.Calculation = tempCalc
    
    Log logLines, "Replaced " & cnt & " integrative activity formulas with values in " & wb.Name & " | Range=" & rng.Address(External:=False)
End Sub

Private Sub ClearIntegrativeActivityZeroGrades(ByVal wb As Object, ByRef logLines As Collection)
    ' Clears only zero values in the integrative activity column (for testing)
    
    If wb Is Nothing Then
        Log logLines, "ERROR: ClearIntegrativeActivityZeroGrades called with Nothing workbook object"
        Exit Sub
    End If
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.Count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.Count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Dim lastRow As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping integrative activity zero grade clearing: " & wb.Name
        Exit Sub
    End If
    
    ' Find the integrative activity column
    Dim integrativeActivityCol As Long
    integrativeActivityCol = GetIntegrativeActivityColumn(ws)
    If integrativeActivityCol = 0 Then
        Log logLines, "INFO: No 'Actividad Integradora' column found in row 2. Skipping integrative activity zero grade clearing: " & wb.Name
        Exit Sub
    End If
    
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, integrativeActivityCol), ws.Cells(lastRow, integrativeActivityCol))
    
    ' Clear cells that evaluate to 0
    Dim cell As Object  ' Late-bound Range
    Dim clearedCount As Long
    Dim cellValue As Variant
    
    For Each cell In rng.Cells
        cellValue = cell.Value
        ' Check if cell value evaluates to 0 (handles 0, 0.0, "0", etc.)
        If IsNumeric(cellValue) And CDbl(cellValue) = 0 Then
            cell.ClearContents
            clearedCount = clearedCount + 1
        End If
    Next cell
    
    Log logLines, "Cleared " & clearedCount & " zero grade cells in integrative activity column for " & wb.Name & " | Range=" & rng.Address(External:=False)
End Sub


