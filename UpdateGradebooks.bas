Attribute VB_Name = "UpdateGradebooks"
Option Explicit

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
    
    Dim fso As Object
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEvents As Boolean
    Dim booEnablePerformanceGuards As Boolean
    
    Dim logLines As Collection
    Set logLines = New Collection
    
    ' Global tracking for all opened workbooks (for error cleanup)
    Dim globalOpenedWorkbooks As Collection
    Set globalOpenedWorkbooks = New Collection
    
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
        OpenMatchingFromSubfolders strBimesterFolderURL, strGradeLevel, openedRefs, globalOpenedWorkbooks, logLines
        
        ' 4.4) Open the template workbook (or use the already-open instance)
        Dim wbTemplate As Workbook
        Set wbTemplate = GetOpenWorkbookByFullPath(fullTemplatePath)
        If wbTemplate Is Nothing Then
            Set wbTemplate = Workbooks.Open(fullTemplatePath)
            ' Track template workbook globally for error cleanup
            globalOpenedWorkbooks.Add fullTemplatePath
            Log logLines, "Opened template: " & fullTemplatePath
        Else
            Log logLines, "Template already open (left as-is): " & fullTemplatePath
        End If
        
        ' 4.5) Replace formulas by values in the single sheet's rectangular range
        On Error Resume Next
        ReplaceFormulasWithValues wbTemplate, logLines
        If Err.Number <> 0 Then
            Log logLines, "ERROR replacing formulas: " & templatePath & " | " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrHandler
        
        ' 4.6) Save & close template
        ' Try to save/close template, but ensure we still close the support files
        Log logLines, "DEBUG: Before template close - Open workbooks count: " & Application.Workbooks.Count
        On Error GoTo TemplateCloseErr
        SafeSaveAndClose wbTemplate, logLines, templatePath
        ' Remove template from global collection since it's now closed
        RemoveFromGlobalCollection globalOpenedWorkbooks, fullTemplatePath
        Log logLines, "DEBUG: After template close - Open workbooks count: " & Application.Workbooks.Count
        On Error GoTo ErrHandler
        
        GoTo AfterTemplate
        
TemplateCloseErr:
        Log logLines, "ERROR while saving/closing template: " & templatePath & " | " & Err.Description
        Err.Clear
        Resume AfterTemplate
        
AfterTemplate:
        ' 4.7) Close only the subfolder files that were opened in step 4.3
        CloseOpenedWorkbooks openedRefs, globalOpenedWorkbooks, logLines
        
NextTemplate:
        templatePath = Dir() ' next .xlsx
    Loop
    
    ' Process completed successfully
    Log logLines, "SUCCESS: Process completed - all workbooks closed properly"
    
    ' Final cleanup check
    FinalCleanupCheck logLines
    
    ' Complete progress bar
    MyProgressbar.Complete
    
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
    MsgBox "GenerateRawGradebooks completed successfully. Check the log for details.", vbInformation, "Process Complete"
    
    Exit Sub

ErrHandler:
    ' Best-effort restore
    On Error Resume Next
    
    ' Close progress bar if it exists
    If Not MyProgressbar Is Nothing Then
        MyProgressbar.Terminate
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

Private Sub ReplaceFormulasWithValues(ByVal wb As Workbook, ByRef logLines As Collection)
    ' Specs:
    ' - Single sheet in template
    ' - Rectangular range starts at C5
    ' - Last row: last non-empty cell in column B
    ' - Last column: last cell in row 3 that has a black background
    ' - Replace only formulas in that area with their current values
    
    Dim ws As Worksheet
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
    
    lastCol = GetLastBlackBackgroundColInRow(ws, 3)  ' row 3 = 3
    If lastCol < 3 Then
        Log logLines, "INFO: No header columns with black background found. Skipping: " & wb.Name
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(5, 3), ws.Cells(lastRow, lastCol)) ' C5 : lastRow,lastCol
    
    ' Replace formulas with values
    Dim cell As Range
    Dim cnt As Long
    For Each cell In rng.Cells
        If cell.HasFormula Then
            cell.value = cell.value
            cnt = cnt + 1
        End If
    Next cell
    
    Log logLines, "Replaced " & cnt & " formulas with values in " & wb.Name & " | Range=" & rng.Address(External:=False)
End Sub

Private Function GetLastNonEmptyRowInColumn(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    Dim lastCell As Range
    Set lastCell = ws.Cells(ws.Rows.Count, colNum).End(xlUp)
    If Len(lastCell.value) = 0 And lastCell.row = 1 Then
        GetLastNonEmptyRowInColumn = 0
    Else
        GetLastNonEmptyRowInColumn = lastCell.row
    End If
End Function

Private Function IsBlackFill(ByVal cell As Range) As Boolean
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

Private Function GetLastBlackBackgroundColInRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
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

                    ' Open only if not already open; track only ones we open now
                    If GetOpenWorkbookByFullPath(fullPath) Is Nothing Then
                        On Error Resume Next
                        Dim wb As Workbook
                        Set wb = Workbooks.Open(fullPath)
                        If Err.Number = 0 Then
                            openedRefs.Add fullPath
                            globalOpenedWorkbooks.Add fullPath  ' Also track globally for error cleanup
                            openedAny = True
                            Log logLines, "Opened: " & fullPath
                            Log logLines, "DEBUG: Added to openedRefs (count=" & openedRefs.Count & ") and globalOpenedWorkbooks (count=" & globalOpenedWorkbooks.Count & ")"
                        Else
                            Log logLines, "ERROR opening: " & fullPath & " | " & Err.Description
                            Err.Clear
                        End If
                        On Error GoTo 0
                    Else
                        ' Already open by the user: do NOT close later
                        Log logLines, "Already open (left as-is): " & fullPath
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
    Dim failedCloses As Collection
    Set failedCloses = New Collection
    
    Log logLines, "DEBUG: Starting to close " & openedRefs.Count & " data files"
    
    For i = openedRefs.Count To 1 Step -1
        Dim p As String
        p = CStr(openedRefs(i))
        
        If Not CloseWorkbookSafely(p, logLines, globalOpenedWorkbooks) Then
            failedCloses.Add p
        End If
        
        openedRefs.Remove i
    Next i
    
    ' Retry failed closes
    If failedCloses.Count > 0 Then
        Log logLines, "DEBUG: Retrying " & failedCloses.Count & " failed closes"
        For i = failedCloses.Count To 1 Step -1
            CloseWorkbookSafely CStr(failedCloses(i)), logLines, globalOpenedWorkbooks
            failedCloses.Remove i
        Next i
    End If
    
    Log logLines, "DEBUG: Finished closing data files"
End Sub

Private Function CloseWorkbookSafely(ByVal path As String, ByRef logLines As Collection, Optional ByRef globalOpenedWorkbooks As Collection = Nothing) As Boolean
    Dim wb As Workbook
    Set wb = GetOpenWorkbookByFullPath(path)
    
    If wb Is Nothing Then
        Log logLines, "DEBUG: Workbook not found (may already be closed): " & path
        CloseWorkbookSafely = True
        Exit Function
    End If
    
    Log logLines, "DEBUG: Attempting to close: " & path
    On Error Resume Next
    wb.Close SaveChanges:=False
    If Err.Number = 0 Then
        ' Verify closure
        Dim wbStillOpen As Workbook
        Set wbStillOpen = GetOpenWorkbookByFullPath(path)
        If wbStillOpen Is Nothing Then
            Log logLines, "Closed: " & path
            ' Remove from global collection if provided
            If Not globalOpenedWorkbooks Is Nothing Then
                RemoveFromGlobalCollection globalOpenedWorkbooks, path
            End If
            CloseWorkbookSafely = True
        Else
            Log logLines, "WARN: Close reported success but workbook still open: " & path
            CloseWorkbookSafely = False
        End If
    Else
        Log logLines, "ERROR closing: " & path & " | " & Err.Description
        CloseWorkbookSafely = False
    End If
    On Error GoTo 0
End Function

Private Sub CloseAllTrackedWorkbooks(ByVal globalOpenedWorkbooks As Collection, ByRef logLines As Collection)
    ' Close all workbooks that were opened during the process (for error cleanup)
    Dim i As Long
    For i = globalOpenedWorkbooks.Count To 1 Step -1
        Dim p As String
        p = CStr(globalOpenedWorkbooks(i))
        
        If CloseWorkbookSafely(p, logLines, globalOpenedWorkbooks) Then
            Log logLines, "ERROR CLEANUP - Closed: " & p
        Else
            Log logLines, "ERROR CLEANUP - Failed to close: " & p
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
    Dim normalizedTarget As String
    Dim normalizedWorkbook As String
    
    ' Normalize the target path
    normalizedTarget = NormalizePath(targetPath)
    
    For Each wb In Application.Workbooks
        On Error Resume Next
        ' Some workbooks may not expose FullName safely; ignore errors
        normalizedWorkbook = NormalizePath(wb.FullName)
        If StrComp(normalizedWorkbook, normalizedTarget, vbTextCompare) = 0 Then
            Set GetOpenWorkbookByFullPath = wb
            Exit Function
        End If
        On Error GoTo 0
    Next wb
    
    Set GetOpenWorkbookByFullPath = Nothing
End Function

Private Function NormalizePath(ByVal path As String) As String
    ' Convert to lowercase and replace backslashes with forward slashes
    Dim normalized As String
    normalized = LCase(Replace(path, "\", "/"))
    
    ' Handle OneDrive vs SharePoint path differences
    If InStr(normalized, "onedrive - abc bilingual school") > 0 Then
        normalized = Replace(normalized, "onedrive - abc bilingual school", "sharepoint.com/personal/jorge_lopez_abcbilingualschool_edu_sv/documents")
    End If
    
    NormalizePath = normalized
End Function

Private Function GetOpenWorkbookByFullPathWithDebug(ByVal targetPath As String, ByRef logLines As Collection) As Workbook
    Dim wb As Workbook
    Dim normalizedTarget As String
    Dim normalizedWorkbook As String
    
    Log logLines, "DEBUG: Looking for workbook with path: " & targetPath
    
    ' Normalize the target path
    normalizedTarget = NormalizePath(targetPath)
    Log logLines, "DEBUG: Normalized target path: " & normalizedTarget
    
    For Each wb In Application.Workbooks
        On Error Resume Next
        ' Some workbooks may not expose FullName safely; ignore errors
        normalizedWorkbook = NormalizePath(wb.FullName)
        Log logLines, "DEBUG: Checking workbook: " & wb.FullName & " -> normalized: " & normalizedWorkbook
        If StrComp(normalizedWorkbook, normalizedTarget, vbTextCompare) = 0 Then
            Log logLines, "DEBUG: MATCH FOUND!"
            Set GetOpenWorkbookByFullPathWithDebug = wb
            Exit Function
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

Private Sub SafeSaveAndClose(ByVal wb As Workbook, ByRef logLines As Collection, Optional ByVal labelName As String = "")
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
' Final Cleanup Check
' ===========================
Private Sub FinalCleanupCheck(ByRef logLines As Collection)
    Dim wb As Workbook
    Dim openCount As Long
    openCount = 0
    
    For Each wb In Application.Workbooks
        ' Skip ThisWorkbook (the current workbook)
        If wb.Name <> ThisWorkbook.Name Then
            openCount = openCount + 1
            Log logLines, "WARNING: Workbook still open: " & wb.FullName
        End If
    Next wb
    
    If openCount > 0 Then
        Log logLines, "FINAL WARNING: " & openCount & " workbooks remain open after process completion"
    Else
        Log logLines, "SUCCESS: All workbooks properly closed"
    End If
End Sub

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


