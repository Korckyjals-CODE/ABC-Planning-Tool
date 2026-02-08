Attribute VB_Name = "TestGradebook"
Option Explicit

' Test version of GenerateRawGradebooks for debugging
' Processes only the first template found with performance guards disabled

Public Sub TestSingleGradebook(ByVal strBimester As String)
    Dim strGradebooksTempFolderURL As String
    Dim strSourceFolderURL As String
    Dim strBimesterFolderURL As String
    
    ' ==== PLACEHOLDERS: set these three before running ====
    strGradebooksTempFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Temp_Grades\")
    strSourceFolderURL = TrimTrailingSlash("C:\Users\korck\OneDrive - ABC BILINGUAL SCHOOL\2526\Computers\Grades\")
    strBimesterFolderURL = JoinPath(strGradebooksTempFolderURL, strBimester)   ' e.g., B1

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
    
    Dim logLines As Collection
    Set logLines = New Collection
    
    ' Global tracking for all opened workbooks (for error cleanup)
    Dim globalOpenedWorkbooks As Collection
    Set globalOpenedWorkbooks = New Collection
    
    ' Track formula placement errors
    Dim formulaErrors As Long
    formulaErrors = 0
    
    On Error GoTo ErrHandler
    
    ' DISABLED for debugging
    booEnablePerformanceGuards = False
    
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
        xlApp.Visible = True  ' MADE VISIBLE FOR DEBUGGING
        xlApp.ScreenUpdating = True  ' ENABLED FOR DEBUGGING
        xlApp.DisplayAlerts = True   ' ENABLED FOR DEBUGGING
        xlApp.Calculation = xlCalculationManual
        xlApp.EnableEvents = False
        Log logLines, "Created VISIBLE Excel instance for debugging"
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
    
    ' 4) Find the FIRST .xlsx template only
    Dim templatePath As String
    templatePath = Dir(strBimesterFolderURL & "\*.xlsx")
    
    If Len(templatePath) = 0 Then
        Log logLines, "ERROR: No .xlsx templates found in " & strBimesterFolderURL
        GoTo Cleanup
    End If
    
    Log logLines, "Found first template: " & templatePath
    
    Dim fullTemplatePath As String
    fullTemplatePath = strBimesterFolderURL & "\" & templatePath
    
    ' 4.1) Extract tag after "Grades-" and before "-Computers"
    Dim strGradeLevelTag As String
    strGradeLevelTag = GetBetween(templatePath, "Grades-", "-Computers")
    
    If Len(strGradeLevelTag) = 0 Then
        Log logLines, "ERROR: Cannot parse grade tag from: " & templatePath
        GoTo Cleanup
    End If
    
    ' 4.2) Map to grade code
    Dim strGradeLevel As String
    strGradeLevel = MapGradeTagToCode(strGradeLevelTag)
    If Len(strGradeLevel) = 0 Then
        Log logLines, "ERROR: Unknown grade mapping for tag '" & strGradeLevelTag & "' in: " & templatePath
        GoTo Cleanup
    End If
    
    Log logLines, "Processing template: " & templatePath & " | Tag='" & strGradeLevelTag & "' -> Code='" & strGradeLevel & "'"
    
    ' 4.3) Open the matching grade workbook(s) from each immediate subfolder
    Dim openedRefs As Collection
    Set openedRefs = New Collection  ' Collection of paths we opened (to close later)
    Log logLines, "DEBUG: About to open matching files for grade: " & strGradeLevel
    OpenMatchingFromSubfolders xlApp, xlAppCreated, strBimesterFolderURL, strGradeLevel, openedRefs, globalOpenedWorkbooks, logLines
    Log logLines, "DEBUG: Opened " & openedRefs.count & " reference files"
    
    ' 4.4) Open the template workbook in the invisible Excel instance
    Dim wbTemplate As Object
    Log logLines, "DEBUG: About to open template: " & fullTemplatePath
    Set wbTemplate = GetOpenWorkbookByFullPathInInstance(xlApp, fullTemplatePath)
    If wbTemplate Is Nothing And xlAppCreated Then
        On Error Resume Next
        Set wbTemplate = xlApp.Workbooks.Open(fullTemplatePath)
        If Err.Number = 0 Then
            ' Track template workbook globally for error cleanup
            globalOpenedWorkbooks.Add fullTemplatePath
            Log logLines, "SUCCESS: Opened template in Excel instance: " & fullTemplatePath
        Else
            Log logLines, "ERROR opening template: " & fullTemplatePath & " | " & Err.Description
            Err.Clear
            Set wbTemplate = Nothing
        End If
        On Error GoTo ErrHandler
    ElseIf Not wbTemplate Is Nothing Then
        Log logLines, "Template already open in Excel instance: " & fullTemplatePath
    Else
        Log logLines, "ERROR: Cannot open template - Excel instance not available"
    End If
    
    ' 4.4.1) Place formula in template if successfully opened
    If Not wbTemplate Is Nothing Then
        Log logLines, "DEBUG: About to place formula in template"
        On Error Resume Next
        PlaceFormulaInTemplate wbTemplate, xlApp, logLines
        If Err.Number <> 0 Then
            formulaErrors = formulaErrors + 1
            Log logLines, "ERROR placing formula in template: " & templatePath & " | " & Err.Description
            Err.Clear
        Else
            Log logLines, "SUCCESS: Formula placed in template"
        End If
        On Error GoTo ErrHandler
    Else
        Log logLines, "SKIP formula placement - template workbook not available"
    End If
    
    ' 4.5) Replace formulas by values in the single sheet's rectangular range
    If Not wbTemplate Is Nothing Then
        Log logLines, "DEBUG: About to replace formulas with values"
        On Error Resume Next
        ReplaceFormulasWithValues wbTemplate, logLines
        If Err.Number <> 0 Then
            Log logLines, "ERROR replacing formulas: " & templatePath & " | " & Err.Description
            Err.Clear
        Else
            Log logLines, "SUCCESS: Formulas replaced with values"
        End If
        On Error GoTo ErrHandler
    Else
        Log logLines, "SKIP replacing formulas - template workbook not available"
    End If
    
    ' 4.6) Save & close template
    If Not wbTemplate Is Nothing Then
        Log logLines, "DEBUG: About to save and close template"
        On Error GoTo TemplateCloseErr
        SafeSaveAndClose wbTemplate, logLines, templatePath
        ' Remove template from global collection since it's now closed
        RemoveFromGlobalCollection globalOpenedWorkbooks, fullTemplatePath
        Log logLines, "SUCCESS: Template saved and closed"
        On Error GoTo ErrHandler
    Else
        Log logLines, "SKIP save/close - template workbook not available"
    End If
    
    GoTo AfterTemplate
    
TemplateCloseErr:
    Log logLines, "ERROR while saving/closing template: " & templatePath & " | " & Err.Description
    Err.Clear
    Resume AfterTemplate
    
AfterTemplate:
    ' 4.7) Close only the subfolder files that were opened in step 4.3
    Log logLines, "DEBUG: About to close reference files"
    CloseOpenedWorkbooks xlApp, xlAppCreated, openedRefs, globalOpenedWorkbooks, logLines
    Log logLines, "SUCCESS: Reference files closed"
    
    ' Process completed successfully
    Log logLines, "SUCCESS: Test process completed"
    
Cleanup:
    ' Clean up the Excel instance
    If xlAppCreated And Not xlApp Is Nothing Then
        On Error Resume Next
        xlApp.Quit
        If Err.Number = 0 Then
            Log logLines, "Closed Excel instance successfully"
        Else
            Log logLines, "WARN: Failed to close Excel instance: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        Set xlApp = Nothing
    End If
    
    ' Flush log
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "Test_GRB_Log"
    
    ' Wrap-up - restore performance guards AFTER logging
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    If prevCalc <> xlCalculationManual Then Application.Calculate
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    
    ' Process completed
    Dim finalMessage As String
    finalMessage = "TestSingleGradebook completed."
    If formulaErrors > 0 Then
        finalMessage = finalMessage & " However, " & formulaErrors & " formula placement error(s) occurred. Check the log for details."
    Else
        finalMessage = finalMessage & " Check the log for details."
    End If
    MsgBox finalMessage, vbInformation, "Test Complete"
    
    Exit Sub

ErrHandler:
    ' Best-effort restore
    On Error Resume Next
    
    ' Close all tracked workbooks before restoring settings
    CloseAllTrackedWorkbooks xlApp, xlAppCreated, globalOpenedWorkbooks, logLines
    
    Log logLines, "FATAL: " & Err.Number & " - " & Err.Description
    DumpLogToImmediate logLines
    DumpLogToSheet logLines, "Test_GRB_Log"
    
    ' Restore performance guards AFTER logging
    Application.Calculation = prevCalc
    If prevCalc <> xlCalculationManual Then Application.Calculate
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    
    MsgBox "TestSingleGradebook encountered an error: " & Err.Description, vbExclamation
End Sub

' ===========================
' Helper Functions (copied from UpdateGradebooks.bas)
' ===========================

Private Sub ReplaceFormulasWithValues(ByVal wb As Object, ByRef logLines As Collection)
    ' Check if workbook object is valid
    If wb Is Nothing Then
        Log logLines, "ERROR: ReplaceFormulasWithValues called with Nothing workbook object"
        Exit Sub
    End If
    
    Log logLines, "DEBUG: ReplaceFormulasWithValues - wb.Name = " & wb.Name
    Log logLines, "DEBUG: ReplaceFormulasWithValues - wb.Worksheets.Count = " & wb.Worksheets.count
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Log logLines, "DEBUG: ReplaceFormulasWithValues - ws.Name = " & ws.Name
    
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    Log logLines, "DEBUG: ReplaceFormulasWithValues - lastRow = " & lastRow
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping: " & wb.Name
        Exit Sub
    End If
    
    lastCol = GetLastBlackBackgroundColInRow(ws, 3)  ' row 3 = 3
    Log logLines, "DEBUG: ReplaceFormulasWithValues - lastCol = " & lastCol
    If lastCol < 3 Then
        Log logLines, "INFO: No header columns with black background found. Skipping: " & wb.Name
        Exit Sub
    End If
    
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    Log logLines, "DEBUG: ReplaceFormulasWithValues - range = " & rng.Address(External:=False)
    
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

Private Sub PlaceFormulaInTemplate(ByVal wb As Object, ByVal xlApp As Object, ByRef logLines As Collection)
    ' Check if workbook object is valid
    If wb Is Nothing Then
        Log logLines, "ERROR: PlaceFormulaInTemplate called with Nothing workbook object"
        Exit Sub
    End If
    
    Log logLines, "DEBUG: PlaceFormulaInTemplate - wb.Name = " & wb.Name
    Log logLines, "DEBUG: PlaceFormulaInTemplate - wb.Worksheets.Count = " & wb.Worksheets.count
    
    Dim ws As Object  ' Late-bound Worksheet
    If wb.Worksheets.count <> 1 Then
        Log logLines, "WARN: Expected 1 sheet, found " & wb.Worksheets.count & " in " & wb.Name & ". Using first sheet."
    End If
    Set ws = wb.Worksheets(1)
    
    Log logLines, "DEBUG: PlaceFormulaInTemplate - ws.Name = " & ws.Name
    
    ' Get the range dimensions
    Dim lastRow As Long, lastCol As Long
    lastRow = GetLastNonEmptyRowInColumn(ws, 2)   ' column B = 2
    Log logLines, "DEBUG: PlaceFormulaInTemplate - lastRow = " & lastRow
    If lastRow < 5 Then
        Log logLines, "INFO: No data rows detected (lastRow < 5). Skipping formula placement: " & wb.Name
        Exit Sub
    End If
    
    lastCol = GetLastWeekColumnInRow(ws, 3)  ' row 3 = 3
    Log logLines, "DEBUG: PlaceFormulaInTemplate - lastCol = " & lastCol
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
    
    Log logLines, "DEBUG: PlaceFormulaInTemplate - formula length = " & Len(formula)
    
    ' Place formula in C5
    On Error Resume Next
    ws.Range("C5").formula = formula
    If Err.Number <> 0 Then
        Log logLines, "ERROR placing formula in " & wb.Name & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Copy formula to the rectangular range
    Dim rng As Object  ' Late-bound Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    
    On Error Resume Next
    ws.Range("C5").Copy
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

' Include all the helper functions from UpdateGradebooks.bas
' (GetLastNonEmptyRowInColumn, IsBlackFill, GetLastBlackBackgroundColInRow, etc.)
' ... (I'll include the essential ones)

Private Function GetLastNonEmptyRowInColumn(ByVal ws As Object, ByVal colNum As Long) As Long
    Dim lastCell As Object  ' Late-bound Range
    Set lastCell = ws.Cells(ws.Rows.count, colNum).End(xlUp)
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
    lastUsedCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).column
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
    lastUsedCol = ws.Cells(rowNum, ws.Columns.count).End(xlToLeft).column
    For c = lastUsedCol To 1 Step -1
        Dim cellValue As String
        cellValue = CStr(ws.Cells(rowNum, c).value)
        If Left(Trim(cellValue), 4) = "Week" Then
            GetLastWeekColumnInRow = c
            Exit Function
        End If
    Next c
    GetLastWeekColumnInRow = 0
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
    Debug.Print "TestSingleGradebook LOG COMPLETE @ " & Now
    Debug.Print String(60, "-")
End Sub

Private Sub DumpLogToSheet(ByVal logLines As Collection, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then
        ws.Cells.Clear
    Else
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If
    On Error GoTo 0
    
    ws.Range("A1").value = "Timestamp"
    ws.Range("B1").value = "Message"
    ws.Range("A1:B1").Font.Bold = True
    
    Dim i As Long
    For i = 1 To logLines.count
        ws.Cells(i + 1, 1).value = Split(logLines(i), "  ")(0)
        ws.Cells(i + 1, 2).value = Mid$(logLines(i), Len(Split(logLines(i), "  ")(0)) + 3)
    Next i
    
    ws.Columns("A:B").EntireColumn.AutoFit
End Sub

' Include the remaining helper functions from UpdateGradebooks.bas
' (OpenMatchingFromSubfolders, CloseOpenedWorkbooks, etc.)
' For brevity, I'll include the essential ones

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
                        On Error Resume Next
                        Dim wb As Object
                        Set wb = xlApp.Workbooks.Open(fullPath)
                        If Err.Number = 0 Then
                            openedRefs.Add fullPath
                            globalOpenedWorkbooks.Add fullPath  ' Also track globally for error cleanup
                            openedAny = True
                            Log logLines, "Opened in Excel instance: " & fullPath
                        Else
                            Log logLines, "ERROR opening in Excel instance: " & fullPath & " | " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    ElseIf Not GetOpenWorkbookByFullPathInInstance(xlApp, fullPath) Is Nothing Then
                        ' Already open in the invisible instance: do NOT close later
                        Log logLines, "Already open in Excel instance: " & fullPath
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
    Log logLines, "DEBUG: Starting to close " & openedRefs.count & " data files"
    For i = openedRefs.count To 1 Step -1
        Dim p As String
        p = CStr(openedRefs(i))
        Log logLines, "DEBUG: Attempting to close data file: " & p
        Dim wb As Object
        If xlAppCreated Then
            Set wb = GetOpenWorkbookByFullPathInInstance(xlApp, p)
            Log logLines, "DEBUG: GetOpenWorkbookByFullPathInInstance returned: " & IIf(wb Is Nothing, "Nothing", "Workbook object")
            If Not wb Is Nothing Then
                Log logLines, "DEBUG: Found open workbook in Excel instance, attempting to close: " & p
                On Error Resume Next
                wb.Close SaveChanges:=False
                If Err.Number = 0 Then
                    ' Verify the workbook is actually closed
                    Dim wbStillOpen As Object
                    Set wbStillOpen = GetOpenWorkbookByFullPathInInstance(xlApp, p)
                    If wbStillOpen Is Nothing Then
                        Log logLines, "Closed in Excel instance: " & p
                        ' Also remove from global collection since it's now closed
                        RemoveFromGlobalCollection globalOpenedWorkbooks, p
                    Else
                        Log logLines, "WARN: Close reported success but workbook still open in Excel instance: " & p
                    End If
                Else
                    Log logLines, "ERROR closing in Excel instance: " & p & " | " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            Else
                Log logLines, "DEBUG: Workbook not found in Excel instance (may already be closed): " & p
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
    For i = globalOpenedWorkbooks.count To 1 Step -1
        Dim p As String
        p = CStr(globalOpenedWorkbooks(i))
        Dim wb As Object
        If xlAppCreated Then
            Set wb = GetOpenWorkbookByFullPathInInstance(xlApp, p)
            If Not wb Is Nothing Then
                On Error Resume Next
                wb.Close SaveChanges:=False
                If Err.Number = 0 Then
                    Log logLines, "ERROR CLEANUP - Closed in Excel instance: " & p
                Else
                    Log logLines, "ERROR CLEANUP - Failed to close in Excel instance: " & p & " | " & Err.Description
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
            Log logLines, "ERROR CLEANUP - Closed Excel instance"
        Else
            Log logLines, "ERROR CLEANUP - Failed to close Excel instance: " & Err.Description
            Err.Clear
        End If
        On Error GoTo 0
        Set xlApp = Nothing
    End If
End Sub

Private Sub RemoveFromGlobalCollection(ByRef globalOpenedWorkbooks As Collection, ByVal pathToRemove As String)
    ' Remove a specific path from the global collection
    Dim i As Long
    For i = globalOpenedWorkbooks.count To 1 Step -1
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
