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
'    - Open matching grade workbook(s) from each subfolder (so formulas can link to open files)
'    - Open template (use open instance if already open)
'    - Place grade lookup formula in C5 and copy to rectangular range (C5 : lastRow, lastCol)
'    - Convert formulas to values in rectangular range (C5 : lastRow, lastCol) per rules
'    - Save & close template
'    - Close only the subfolder files we opened for this template
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
Public Sub GenerateRawGradebooks(ByVal strBimester As String)
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
    
    ' New invisible Excel instance for opening files
    Dim xlApp As Object
    Dim xlAppCreated As Boolean
    
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
    
    ' Create invisible Excel instance for opening files
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    If Err.Number = 0 Then
        xlAppCreated = True
        xlApp.Visible = False
        xlApp.ScreenUpdating = False
        xlApp.DisplayAlerts = False
        xlApp.Calculation = xlCalculationManual
        xlApp.EnableEvents = False
        Log logLines, "Created invisible Excel instance for file operations"
    Else
        Log logLines, "ERROR: Failed to create Excel instance - " & Err.Description
        Err.Clear
        xlAppCreated = False
    End If
    On Error GoTo ErrHandler
    
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
        
        Log logLines, "Processing template: " & templatePath & " | Tag='" & strGradeLevelTag & "' ? Code='" & strGradeLevel & "'"
        
        ' Update progress bar
        MyProgressbar.NextAction "Processing '" & templatePath & "'", True
        
        ' 4.3) Open the matching grade workbook(s) from each immediate subfolder
        Dim openedRefs As Collection
        Set openedRefs = New Collection  ' Collection of paths we opened (to close later)
        OpenMatchingFromSubfolders xlApp, xlAppCreated, strBimesterFolderURL, strGradeLevel, openedRefs, globalOpenedWorkbooks, logLines
        
        ' 4.4) Open the template workbook in the invisible Excel instance
        Dim wbTemplate As Object
        Set wbTemplate = GetOpenWorkbookByFullPathInInstance(xlApp, fullTemplatePath)
        If wbTemplate Is Nothing And xlAppCreated Then
            ' Use watchdog approach to handle COM object timing issues
            Set wbTemplate = OpenWorkbookWithWatchdog(xlApp, fullTemplatePath, 10, logLines)
            If Not wbTemplate Is Nothing Then
                ' Track template workbook globally for error cleanup
                globalOpenedWorkbooks.Add fullTemplatePath
            End If
        ElseIf Not wbTemplate Is Nothing Then
            Log logLines, "Template already open in invisible instance: " & fullTemplatePath
        Else
            Log logLines, "ERROR: Cannot open template - Excel instance not available"
        End If
        
        ' 4.4.1) Place formula in template if successfully opened
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            PlaceFormulaInTemplate wbTemplate, xlApp, logLines
            If Err.Number <> 0 Then
                formulaErrors = formulaErrors + 1
                Log logLines, "ERROR placing formula in template: " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.5) Replace formulas by values in the single sheet's rectangular range
        If Not wbTemplate Is Nothing Then
            ' More robust validation: check if workbook is still valid
            On Error Resume Next
            Dim testCount As Long
            testCount = wbTemplate.Worksheets.Count  ' This is safer than accessing .Name
            If Err.Number <> 0 Then
                Log logLines, "WARN: Workbook object became invalid before ReplaceFormulasWithValues: " & templatePath
                Err.Clear
                Set wbTemplate = Nothing
            End If
            On Error GoTo ErrHandler
            
            If Not wbTemplate Is Nothing Then
                On Error Resume Next
                ReplaceFormulasWithValues wbTemplate, logLines
                If Err.Number <> 0 Then
                    Log logLines, "ERROR replacing formulas: " & templatePath & " | " & Err.Description
                    Err.Clear
                End If
                On Error GoTo ErrHandler
            Else
                Log logLines, "SKIP replacing formulas - template workbook became invalid: " & templatePath
            End If
        Else
            Log logLines, "SKIP replacing formulas - template workbook not available: " & templatePath
        End If
        
        ' 4.6) Run health check on generated gradebook (optional)
        If Not wbTemplate Is Nothing Then
            On Error Resume Next
            RunHealthCheckOnGeneratedGradebook wbTemplate, templatePath
            If Err.Number <> 0 Then
                Log logLines, "WARN: Health check failed for " & templatePath & " | " & Err.Description
                Err.Clear
            End If
            On Error GoTo ErrHandler
        End If
        
        ' 4.7) Save & close template
        ' Try to save/close template, but ensure we still close the support files
        Log logLines, "DEBUG: Before template close - Open workbooks count: " & Application.Workbooks.Count
        If Not wbTemplate Is Nothing Then
            On Error GoTo TemplateCloseErr
            SafeSaveAndClose wbTemplate, logLines, templatePath
            ' Remove template from global collection since it's now closed
            RemoveFromGlobalCollection globalOpenedWorkbooks, fullTemplatePath
            Log logLines, "DEBUG: After template close - Open workbooks count: " & Application.Workbooks.Count
            On Error GoTo ErrHandler
        Else
            Log logLines, "SKIP save/close - template workbook not available: " & templatePath
        End If
        
        GoTo AfterTemplate
        
TemplateCloseErr:
        Log logLines, "ERROR while saving/closing template: " & templatePath & " | " & Err.Description
        Err.Clear
        Resume AfterTemplate
        
AfterTemplate:
        ' 4.7) Close only the subfolder files that were opened in step 4.3
        CloseOpenedWorkbooks xlApp, xlAppCreated, openedRefs, globalOpenedWorkbooks, logLines
        
NextTemplate:
        templatePath = Dir() ' next .xlsx
    Loop
    
    ' Process completed successfully
    Log logLines, "SUCCESS: Process completed - all workbooks closed properly"
    
    ' Clean up the invisible Excel instance
    If xlAppCreated And Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        If Err.Number = 0 Then
            Log logLines, "Closed invisible Excel instance successfully"
        Else
            Log logLines, "WARN: Failed to close invisible Excel instance: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        Set xlApp = Nothing
    End If
    
    ' Complete progress bar
    MyProgressbar.Complete
    
    ' Close progress bar after completion
    MyProgressbar.Terminate
    Set MyProgressbar = Nothing
    
    ' Flush log (while performance guards are still active)
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "GRB_Log"
    
    ' Wrap-up - restore performance guards AFTER logging
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    If prevCalc <> xlCalculationManual Then Application.Calculate
    
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
    MsgBox finalMessage, vbInformation, "Process Complete"
    
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
    CloseAllTrackedWorkbooks xlApp, xlAppCreated, globalOpenedWorkbooks, logLines
    
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
    
    ' Check if workbook object is valid
    If wb Is Nothing Then
        Log logLines, "ERROR: ReplaceFormulasWithValues called with Nothing workbook object"
        Exit Sub
    End If
    
    ' Additional validation - check if workbook is still accessible
    On Error Resume Next
    Dim testCount As Long
    testCount = wb.Worksheets.Count
    If Err.Number <> 0 Then
        Log logLines, "ERROR: Workbook object is no longer accessible in ReplaceFormulasWithValues"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
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
    
    ' Replace formulas with values
    Dim cell As Object  ' Late-bound Range
    Dim cnt As Long
    For Each cell In rng.Cells
        If cell.HasFormula Then
            cell.value = cell.value
            cnt = cnt + 1
        End If
    Next cell
    
    Log logLines, "Replaced " & cnt & " formulas with values in " & wb.Name & " | Range=" & rng.Address(External:=False)
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

Private Sub PlaceFormulaInTemplate(ByVal wb As Object, ByVal xlApp As Object, ByRef logLines As Collection)
    ' Places the grade lookup formula in the template and copies it to the appropriate range
    
    ' Check if workbook object is valid
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
    formula = formula & """Ciclo Tres Development Center_B"",""DC3B"")),"
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
    formula = formula & "grade,IFERROR(XLOOKUP(clean_name,name_rng_xlsx,grade_rng_xlsx,""""),XLOOKUP(clean_name,name_rng_xlsm,grade_rng_xlsm,"""")),"
    formula = formula & "IFERROR(grade,""""))"
    
    ' Place formula in C5 using watchdog approach
    If Not PlaceFormulaWithWatchdog(ws, formula, 10, logLines) Then
        Log logLines, "ERROR: Failed to place formula in " & wb.Name & " after timeout"
        Exit Sub
    End If
    
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
    
    ' Clear clipboard using the correct Excel instance
    If Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.CutCopyMode = False
        On Error GoTo 0
    End If
End Sub

Private Sub OpenMatchingFromSubfolders(ByVal xlApp As Object, ByVal xlAppCreated As Boolean, ByVal bimesterFolder As String, ByVal gradeCode As String, _
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

                    ' Open only if not already open in the invisible instance; track only ones we open now
                    If GetOpenWorkbookByFullPathInInstance(xlApp, fullPath) Is Nothing And xlAppCreated Then
                        ' Use watchdog approach to handle COM object timing issues
                        Dim wb As Object
                        Set wb = OpenWorkbookWithWatchdog(xlApp, fullPath, 10, logLines)
                        If Not wb Is Nothing Then
                            openedRefs.Add fullPath
                            globalOpenedWorkbooks.Add fullPath  ' Also track globally for error cleanup
                            openedAny = True
                            Log logLines, "DEBUG: Added to openedRefs (count=" & openedRefs.Count & ") and globalOpenedWorkbooks (count=" & globalOpenedWorkbooks.Count & ")"
                        End If
                    ElseIf Not GetOpenWorkbookByFullPathInInstance(xlApp, fullPath) Is Nothing Then
                        ' Already open in the invisible instance: do NOT close later
                        Log logLines, "Already open in invisible instance: " & fullPath
                    Else
                        Log logLines, "ERROR: Cannot open - Excel instance not available: " & fullPath
                    End If
                End If
            End If
        Next fil

        If Not openedAny Then
            Log logLines, "No matching file in subfolder: " & subf.path & " (pattern '" & pattern & "')"
        End If
    Next subf
End Sub

Private Sub CloseOpenedWorkbooks(ByVal xlApp As Object, ByVal xlAppCreated As Boolean, ByVal openedRefs As Collection, ByRef globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    Dim i As Long
    Log logLines, "DEBUG: Starting to close " & openedRefs.Count & " data files"
    For i = openedRefs.Count To 1 Step -1
        Dim p As String
        p = CStr(openedRefs(i))
        Log logLines, "DEBUG: Attempting to close data file: " & p
        Dim wb As Object
        If xlAppCreated Then
            Set wb = GetOpenWorkbookByFullPathInInstance(xlApp, p)
            Log logLines, "DEBUG: GetOpenWorkbookByFullPathInInstance returned: " & IIf(wb Is Nothing, "Nothing", "Workbook object")
            If Not wb Is Nothing Then
                Log logLines, "DEBUG: Found open workbook in invisible instance, attempting to close: " & p
                On Error Resume Next
                wb.Close SaveChanges:=False
                If Err.Number = 0 Then
                    ' Verify the workbook is actually closed
                    Dim wbStillOpen As Object
                    Set wbStillOpen = GetOpenWorkbookByFullPathInInstance(xlApp, p)
                    If wbStillOpen Is Nothing Then
                        Log logLines, "Closed in invisible instance: " & p
                        ' Also remove from global collection since it's now closed
                        RemoveFromGlobalCollection globalOpenedWorkbooks, p
                    Else
                        Log logLines, "WARN: Close reported success but workbook still open in invisible instance: " & p
                        ' Try more aggressive closing
                        On Error Resume Next
                        wbStillOpen.Close SaveChanges:=False
                        If Err.Number = 0 Then
                            Log logLines, "Retry close successful in invisible instance: " & p
                            RemoveFromGlobalCollection globalOpenedWorkbooks, p
                        Else
                            Log logLines, "Retry close failed in invisible instance: " & p & " | " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                Else
                    Log logLines, "ERROR closing in invisible instance: " & p & " | " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                Log logLines, "DEBUG: Workbook not found in invisible instance (may already be closed): " & p
            End If
        Else
            Log logLines, "ERROR: Cannot close workbook - Excel instance not available: " & p
        End If
        openedRefs.Remove i
    Next i
    Log logLines, "DEBUG: Finished closing data files"
End Sub

Private Sub CloseAllTrackedWorkbooks(ByVal xlApp As Object, ByVal xlAppCreated As Boolean, ByVal globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    ' Close all workbooks that were opened during the process (for error cleanup)
    Dim i As Long
    For i = globalOpenedWorkbooks.Count To 1 Step -1
        Dim p As String
        p = CStr(globalOpenedWorkbooks(i))
        Dim wb As Object
        If xlAppCreated Then
            Set wb = GetOpenWorkbookByFullPathInInstance(xlApp, p)
            If Not wb Is Nothing Then
                On Error Resume Next
                wb.Close SaveChanges:=False
                If Err.Number = 0 Then
                    ' Verify the workbook is actually closed
                    Dim wbStillOpen As Object
                    Set wbStillOpen = GetOpenWorkbookByFullPathInInstance(xlApp, p)
                    If wbStillOpen Is Nothing Then
                        Log logLines, "ERROR CLEANUP - Closed in invisible instance: " & p
                    Else
                        Log logLines, "ERROR CLEANUP - WARN: Close reported success but workbook still open in invisible instance: " & p
                    End If
                Else
                    Log logLines, "ERROR CLEANUP - Failed to close in invisible instance: " & p & " | " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
        globalOpenedWorkbooks.Remove i
    Next i
    
    ' Clean up the Excel instance
    If xlAppCreated And Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        If Err.Number = 0 Then
            Log logLines, "ERROR CLEANUP - Closed invisible Excel instance"
        Else
            Log logLines, "ERROR CLEANUP - Failed to close invisible Excel instance: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        Set xlApp = Nothing
    End If
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

Private Function GetOpenWorkbookByFullPathInInstance(ByVal xlApp As Object, ByVal targetPath As String) As Object
    Dim wb As Object
    Dim localPath As String
    Dim sharePointPath As String
    
    ' Try direct match first
    For Each wb In xlApp.Workbooks
        On Error Resume Next
        ' Some workbooks may not expose FullName safely; ignore errors
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullPathInInstance = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
    
    ' If no direct match, try to convert local OneDrive path to SharePoint URL
    If InStr(targetPath, "OneDrive - ABC BILINGUAL SCHOOL") > 0 Then
        localPath = targetPath
        sharePointPath = ConvertLocalPathToSharePointURL(localPath)
        
        For Each wb In xlApp.Workbooks
            On Error Resume Next
            If StrComp(wb.FullName, sharePointPath, vbTextCompare) = 0 Then
                Set GetOpenWorkbookByFullPathInInstance = wb
                Exit Function
            End If
            On Error GoTo 0
        Next wb
    End If
    
    ' If still no match, try to convert SharePoint URL to local path
    For Each wb In xlApp.Workbooks
        On Error Resume Next
        If InStr(wb.FullName, "sharepoint.com") > 0 Then
            localPath = ConvertSharePointURLToLocalPath(wb.FullName)
            If StrComp(localPath, targetPath, vbTextCompare) = 0 Then
                Set GetOpenWorkbookByFullPathInInstance = wb
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next wb
    
    Set GetOpenWorkbookByFullPathInInstance = Nothing
End Function

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

Private Function GetOpenWorkbookByFullPathWithDebug(ByVal targetPath As String, ByRef logLines As Collection) As Workbook
    Dim wb As Workbook
    Dim localPath As String
    Dim sharePointPath As String
    
    Log logLines, "DEBUG: Looking for workbook with path: " & targetPath
    
    ' Try direct match first
    For Each wb In Application.Workbooks
        On Error Resume Next
        ' Some workbooks may not expose FullName safely; ignore errors
        Log logLines, "DEBUG: Checking workbook: " & wb.FullName
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Log logLines, "DEBUG: MATCH FOUND!"
            Set GetOpenWorkbookByFullPathWithDebug = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
    
    ' If no direct match, try to convert local OneDrive path to SharePoint URL
    If InStr(targetPath, "OneDrive - ABC BILINGUAL SCHOOL") > 0 Then
        localPath = targetPath
        sharePointPath = ConvertLocalPathToSharePointURL(localPath)
        Log logLines, "DEBUG: Converted to SharePoint URL: " & sharePointPath
        
        For Each wb In Application.Workbooks
            On Error Resume Next
            Log logLines, "DEBUG: Checking workbook: " & wb.FullName
            If StrComp(wb.FullName, sharePointPath, vbTextCompare) = 0 Then
                Log logLines, "DEBUG: MATCH FOUND with converted path!"
                Set GetOpenWorkbookByFullPathWithDebug = wb
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
            Log logLines, "DEBUG: Converted SharePoint URL to local path: " & localPath
            If StrComp(localPath, targetPath, vbTextCompare) = 0 Then
                Log logLines, "DEBUG: MATCH FOUND with reverse conversion!"
                Set GetOpenWorkbookByFullPathWithDebug = wb
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next wb
    
    Log logLines, "DEBUG: No match found for: " & targetPath
    Set GetOpenWorkbookByFullPathWithDebug = Nothing
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

Private Function GetBetween(ByVal text As String, ByVal after As String, ByVal before As String) As String
    Dim p1 As Long, p2 As Long, startPos As Long
    p1 = InStr(1, text, after, vbTextCompare)
    If p1 = 0 Then Exit Function
    startPos = p1 + Len(after)
    p2 = InStr(startPos, text, before, vbTextCompare)
    If p2 = 0 Then Exit Function
    GetBetween = Mid$(text, startPos, p2 - startPos)
End Function

Private Function MapGradeTagToCode(ByVal tag As String) As String
    ' Handles:
    '  - "First Grade_A" ? "1A", "First Grade_B" ? "1B", ..., "Twelfth Grade_A/B" ? "12A/12B"
    '  - "Ciclo Tres Development Center_A" ? "DC3A", "_B" ? "DC3B"
    Dim sec As String
    sec = Right$(tag, 2) ' expects "_A" or "_B"
    If Not (sec = "_A" Or sec = "_B") Then
        MapGradeTagToCode = ""
        Exit Function
    End If
    
    Dim suffix As String
    suffix = Right$(sec, 1) ' "A" or "B"
    
    ' Special DC3
    If InStr(1, tag, "Ciclo Tres Development Center", vbTextCompare) = 1 Then
        MapGradeTagToCode = "DC3" & suffix
        Exit Function
    End If
    
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

Private Sub SafeSaveAndClose(ByVal wb As Object, ByRef logLines As Collection, Optional ByVal labelName As String = "")
    Dim i As Long
    Dim okSave As Boolean, okClose As Boolean
    
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        wb.Save
        okSave = (Err.Number = 0)
        If Not okSave Then
            Log logLines, "WARN: Save retry " & i & " for " & IIf(Len(labelName) > 0, labelName, wb.Name) & " | " & Err.Description
            DoEvents
            SleepShort 300
        Else
            Log logLines, "DEBUG: Successfully saved: " & IIf(Len(labelName) > 0, labelName, wb.Name)
            Exit For
        End If
    Next i
    
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        wb.Close SaveChanges:=False
        okClose = (Err.Number = 0)
        If Not okClose Then
            Log logLines, "WARN: Close retry " & i & " for " & IIf(Len(labelName) > 0, labelName, wb.Name) & " | " & Err.Description
            DoEvents
            SleepShort 300
        Else
            Log logLines, "DEBUG: Successfully closed: " & IIf(Len(labelName) > 0, labelName, wb.Name)
            Exit For
        End If
    Next i
    
    On Error GoTo 0
End Sub

Private Sub SleepShort(ms As Long)
    Dim t As Single: t = Timer
    Do While Timer - t < ms / 1000!
        DoEvents
    Loop
End Sub

' ===========================
' Watchdog Functions for Timing Issues
' ===========================

Private Function OpenWorkbookWithWatchdog(ByVal xlApp As Object, ByVal filePath As String, ByVal timeoutSeconds As Long, ByRef logLines As Collection) As Object
    ' Opens a workbook with retry logic and timeout to handle COM object timing issues
    Dim startTime As Double
    Dim wb As Object
    Dim success As Boolean
    Dim attemptCount As Long
    
    startTime = Timer
    success = False
    attemptCount = 0
    
    Do While (Timer - startTime) < timeoutSeconds
        attemptCount = attemptCount + 1
        
        On Error Resume Next
        Set wb = xlApp.Workbooks.Open(filePath)
        If Err.Number = 0 And Not wb Is Nothing Then
            ' Verify the workbook is actually accessible
            Dim testCount As Long
            testCount = wb.Worksheets.Count
            If Err.Number = 0 Then
                success = True
                Log logLines, "SUCCESS: Opened workbook (attempt " & attemptCount & "): " & filePath
                Exit Do
            Else
                Log logLines, "WARN: Workbook opened but not accessible (attempt " & attemptCount & "): " & filePath
                Err.Clear
                If Not wb Is Nothing Then
                    wb.Close SaveChanges:=False
                    Set wb = Nothing
                End If
            End If
        Else
            If attemptCount = 1 Then
                Log logLines, "WARN: Failed to open workbook (attempt " & attemptCount & "): " & filePath & " | " & Err.Description
            End If
            Err.Clear
        End If
        On Error GoTo 0
        
        ' Small delay before retry
        DoEvents
        SleepShort 50
    Loop
    
    If success Then
        Set OpenWorkbookWithWatchdog = wb
    Else
        Log logLines, "ERROR: Failed to open workbook after " & attemptCount & " attempts in " & Format$(Timer - startTime, "0.0") & " seconds: " & filePath
        Set OpenWorkbookWithWatchdog = Nothing
    End If
End Function

Private Function PlaceFormulaWithWatchdog(ByVal ws As Object, ByVal formula As String, ByVal timeoutSeconds As Long, ByRef logLines As Collection) As Boolean
    ' Places a formula with retry logic and timeout to handle worksheet timing issues
    Dim startTime As Double
    Dim success As Boolean
    Dim attemptCount As Long
    
    startTime = Timer
    success = False
    attemptCount = 0
    
    Do While (Timer - startTime) < timeoutSeconds
        attemptCount = attemptCount + 1
        
        On Error Resume Next
        ws.Range(FORMULA_START_CELL).Formula = formula
        If Err.Number = 0 Then
            success = True
            Log logLines, "SUCCESS: Placed formula (attempt " & attemptCount & ") in " & ws.Parent.Name
            Exit Do
        Else
            If attemptCount = 1 Then
                Log logLines, "WARN: Failed to place formula (attempt " & attemptCount & ") in " & ws.Parent.Name & " | " & Err.Description
            End If
            Err.Clear
        End If
        On Error GoTo 0
        
        ' Small delay before retry
        DoEvents
        SleepShort 50
    Loop
    
    If Not success Then
        Log logLines, "ERROR: Failed to place formula after " & attemptCount & " attempts in " & Format$(Timer - startTime, "0.0") & " seconds in " & ws.Parent.Name
    End If
    
    PlaceFormulaWithWatchdog = success
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
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
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


