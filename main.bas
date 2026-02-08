Attribute VB_Name = "main"
Option Explicit


Function DurationCodeToMins(ByVal strDurationCode As String, Optional ByVal intClassDuration As Integer = 45) As Double

Dim strSuffix As String
Dim strDuration As String
Dim dblDuration As Double

If strDurationCode = "" Then
    dblDuration = 0
Else
    strSuffix = Right(strDurationCode, 1)
    strDuration = Left(strDurationCode, Len(strDurationCode) - 1)
    Select Case strSuffix
        Case "m"
            dblDuration = CDbl(strDuration)
        Case "c"
            dblDuration = CDbl(strDuration) * intClassDuration
    End Select
End If

DurationCodeToMins = dblDuration

End Function

Function ExpandArrays(DatesRange As Range, StringsRange As Range) As Variant
    Dim i As Long, j As Long, k As Long
    Dim SplitItems() As String
    Dim ResultArray() As Variant
    Dim TotalRows As Long
    Dim DatesArray As Variant
    Dim StringsArray As Variant

    ' Assign the range values to arrays
    DatesArray = DatesRange.value
    StringsArray = StringsRange.value

    ' Calculate the total number of rows needed for the output array
    TotalRows = 0
    For i = LBound(StringsArray, 1) To UBound(StringsArray, 1)
        SplitItems = Split(StringsArray(i, 1), ", ")
        TotalRows = TotalRows + UBound(SplitItems) - LBound(SplitItems) + 1
    Next i

    ' Redimension the result array to hold the output
    ReDim ResultArray(1 To TotalRows, 1 To 2)

    ' Fill the result array with dates and corresponding items
    k = 1
    For i = LBound(DatesArray, 1) To UBound(DatesArray, 1)
        SplitItems = Split(StringsArray(i, 1), ", ")
        For j = LBound(SplitItems) To UBound(SplitItems)
            ResultArray(k, 1) = DatesArray(i, 1)
            ResultArray(k, 2) = SplitItems(j)
            k = k + 1
        Next j
    Next i

    ' Return the result array
    ExpandArrays = ResultArray
End Function

Sub GenerateNextAgenda()

Dim tbl As ListObject
Dim ws As Worksheet
Dim i As Integer
Dim strGrade As String
Dim strPlannedActivity As String
Dim datPlannedDate As Date
Dim intLastRowIndex As Integer
Dim intLastColIndex As Integer
Dim datTargetDate As Date
Dim strAgendaLine As String
Dim intTargetDayNumber As Integer
Dim strTargetDateName As String
Dim tblDays As ListObject
Dim rngTargetDay As Range
Dim tblSchedule As ListObject
Dim rngCell As Range
Dim strWsCodename As String
Dim strAgenda As String

Set tblDays = wsDays.ListObjects("tblDays")

datTargetDate = ShowDatePicker()

intTargetDayNumber = Weekday(datTargetDate)

strTargetDateName = Application.WorksheetFunction.XLookup(intTargetDayNumber, _
    tblDays.ListColumns(1).DataBodyRange, tblDays.ListColumns(2).DataBodyRange)
    
Set tblSchedule = wsSchedule.ListObjects("tblSchedule")

Set rngTargetDay = tblSchedule.HeaderRowRange.Find(What:=strTargetDateName, LookIn:=XlFindLookIn.xlValues, _
    LookAt:=XlLookAt.xlWhole)

Set rngCell = rngTargetDay.Offset(1, 0)
Do While rngCell.value <> ""
    If HasDigit(rngCell.value) Then
        strGrade = Replace(Replace(rngCell.value, "A", ""), "B", "")
        strWsCodename = "ws" & strGrade & "ActTBox"
        Set ws = GetWorksheetByCodename(strWsCodename)
        For Each tbl In ws.ListObjects
            strGrade = tbl.HeaderRowRange(1, 2).Offset(-1, 0)
            intLastRowIndex = tbl.ListRows.count
            intLastColIndex = tbl.ListColumns.count
            strPlannedActivity = tbl.ListColumns(2).DataBodyRange(intLastRowIndex)
            datPlannedDate = tbl.ListColumns(intLastColIndex - 1).DataBodyRange(intLastRowIndex)
            If datPlannedDate = datTargetDate Then
                strAgendaLine = strGrade & ": " & strPlannedActivity
                Debug.Print strAgendaLine
                strAgenda = strAgenda & strAgendaLine & vbCrLf
            End If
        Next
    End If
    Set rngCell = rngCell.Offset(1, 0)
Loop


End Sub

Sub Generate()



End Sub

Function GetActivityList(ByVal strStartActivity As String, ByVal strSection As String, ByVal strBlock As String) As Variant

Dim colActivities As New Collection
Dim tblTarget As ListObject
Dim intBlockDuration As Integer
Dim rngCurrentActivity As Range
Dim intActivityDuration As Integer
Dim intTotalDuration As Integer
Dim colObjective As New Collection
Dim colDescription As New Collection

Set tblTarget = GetMasterTable(strSection)
intBlockDuration = GetBlockDuration(strBlock)

Set rngCurrentActivity = tblTarget.ListColumns(2).DataBodyRange.Find(What:=strStartActivity, LookIn:=Excel.XlFindLookIn.xlValues, _
    LookAt:=Excel.XlLookAt.xlWhole)
intTotalDuration = 0
Do While Not rngCurrentActivity Is Nothing
    colActivities.Add rngCurrentActivity.value
    colObjective.Add rngCurrentActivity.Offset(0, 2).value
    colDescription.Add rngCurrentActivity.Offset(0, 3).value
    intActivityDuration = Int(DurationCodeToMins(rngCurrentActivity.Offset(0, 1).value))
    intTotalDuration = intTotalDuration + intActivityDuration
    If intTotalDuration >= intBlockDuration Then
        Set rngCurrentActivity = Nothing
    Else
        Set rngCurrentActivity = rngCurrentActivity.Offset(1, 0)
    End If
Loop

GetActivityList = Array(colActivities, colObjective, colDescription)

End Function

Function GetBimonthly(ByVal strBimester As String) As String

Dim tbl As ListObject

Set tbl = wsBimesters.ListObjects("tblBimesterToBimonthly")

strBimester = tbl.Application.WorksheetFunction.XLookup(strBimester, tbl.ListColumns(1).DataBodyRange, tbl.ListColumns(2).DataBodyRange)

GetBimonthly = strBimester

End Function

Function GetDateRange(ByVal strBimester As String, ByVal strSection As String) As String
    ' Get the Date Range (column 7) for the corresponding Bimester (column 2) and Section (column 1),
    ' from tblProjects in sheet wsProjects.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsProjects.ListObjects("tblProjects")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetDateRange = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 1)) = strSection And Trim(data(i, 2)) = strBimester Then
            GetDateRange = data(i, 7)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetDateRange = ""
End Function


Function GetGradeNumberFromSection(ByVal strSection As String) As String
    ' Returns the string before the uppercase letter A or B at the end of strSection.
    ' Example: 1A returns 1.
    ' Example: DC3B returns DC3.

    Dim lastChar As String
    Dim sectionLength As Long

    sectionLength = Len(strSection)

    If sectionLength = 0 Then
        GetGradeNumberFromSection = ""
        Exit Function
    End If

    lastChar = Right(strSection, 1)

    If lastChar = "A" Or lastChar = "B" Then
        GetGradeNumberFromSection = Left(strSection, sectionLength - 1)
    Else
        GetGradeNumberFromSection = strSection
    End If
End Function


Function GetProjectNumberForUnitPlan(ByVal strBimester As String, ByVal strSection As String) As String
    Dim tbl As ListObject
    Dim i As Long
    Dim cellValue As String
    Dim regex As Object
    Dim matches As Object

    ' Set the worksheet and table
    Set tbl = wsProjects.ListObjects("tblProjects")

    ' Initialize RegExp to find "Project N"
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "Project\s+(\d+)"
    regex.IgnoreCase = True
    regex.Global = False

    ' Loop through the rows in the table
    For i = 1 To tbl.ListRows.count
        With tbl.ListRows(i).Range
            If Trim(.Cells(1, 2).value) = strBimester And Trim(.Cells(1, 1).value) = strSection Then
                cellValue = .Cells(1, 3).value ' Column 3 has the project title
                If regex.test(cellValue) Then
                    Set matches = regex.Execute(cellValue)
                    GetProjectNumberForUnitPlan = matches(0).SubMatches(0)
                    Exit Function
                End If
            End If
        End With
    Next i

    ' If not found
    GetProjectNumberForUnitPlan = ""
End Function


Function GetUnitPlanModules(ByVal intProjectNumber As Integer, ByVal strGrade As String) As String
    ' Get the Activities (column 6) for the corresponding Project # (column 1) and Grade (column 3),
    ' from tblTBoxProjectsInfo in sheet wsTBoxProjectsInfo.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsTBoxProjectsInfo.ListObjects("tblTBoxProjectsInfo")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetUnitPlanModules = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 3)) = strGrade And Trim(data(i, 1)) = intProjectNumber Then
            GetUnitPlanModules = data(i, 6)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetUnitPlanModules = ""
End Function

Function GetUnitPlanStandards(ByVal intProjectNumber As Integer, ByVal strGrade As String) As String
    ' Get the Standards (column 7) for the corresponding Project # (column 1) and Grade (column 3),
    ' from tblTBoxProjectsInfo in sheet wsTBoxProjectsInfo.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsTBoxProjectsInfo.ListObjects("tblTBoxProjectsInfo")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetUnitPlanStandards = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 3)) = strGrade And Trim(data(i, 1)) = intProjectNumber Then
            GetUnitPlanStandards = data(i, 7)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetUnitPlanStandards = ""
End Function

Function GetUnitPlanContents(ByVal intProjectNumber As Integer, ByVal strGrade As String) As String
    ' Get the Contents (column 8) for the corresponding Project # (column 1) and Grade (column 3),
    ' from tblTBoxProjectsInfo in sheet wsTBoxProjectsInfo.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsTBoxProjectsInfo.ListObjects("tblTBoxProjectsInfo")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetUnitPlanContents = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 3)) = strGrade And Trim(data(i, 1)) = intProjectNumber Then
            GetUnitPlanContents = data(i, 8)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetUnitPlanContents = ""
End Function


Function GetUnitPlanObjectives(ByVal intProjectNumber As Integer, ByVal strGrade As String) As String
    ' Get the Objectives (column 9) for the corresponding Project # (column 1) and Grade (column 3),
    ' from tblTBoxProjectsInfo in sheet wsTBoxProjectsInfo.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsTBoxProjectsInfo.ListObjects("tblTBoxProjectsInfo")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetUnitPlanObjectives = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 3)) = strGrade And Trim(data(i, 1)) = intProjectNumber Then
            GetUnitPlanObjectives = data(i, 9)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetUnitPlanObjectives = ""
End Function

Function GetUnitPlanIndicators(ByVal intProjectNumber As Integer, ByVal strGrade As String) As String
    ' Get the Indicators (column 10) for the corresponding Project # (column 1) and Grade (column 3),
    ' from tblTBoxProjectsInfo in sheet wsTBoxProjectsInfo.

    Dim tbl As ListObject
    Dim i As Long
    Dim data As Variant

    Set tbl = wsTBoxProjectsInfo.ListObjects("tblTBoxProjectsInfo")

    ' Exit if table has no data
    If tbl.DataBodyRange Is Nothing Then
        GetUnitPlanIndicators = ""
        Exit Function
    End If

    ' Read table data into an array (faster than looping cells)
    data = tbl.DataBodyRange.value

    For i = 1 To UBound(data, 1)
        If Trim(data(i, 3)) = strGrade And Trim(data(i, 1)) = intProjectNumber Then
            GetUnitPlanIndicators = data(i, 10)
            Exit Function
        End If
    Next i

    ' If no match is found
    GetUnitPlanIndicators = ""
End Function

Function GetNextActivity(ByVal strStartActivity As String, ByVal intTargetTotalDuration As Integer, ByVal strSection) As String

Dim colActivities As New Collection
Dim tblTarget As ListObject
Dim rngCurrentActivity As Range
Dim intActivityDuration As Integer
Dim intTotalDuration As Integer
Dim colObjective As New Collection
Dim colDescription As New Collection
Dim rngNextActivity As Range
Dim strNextActivity As String

Set tblTarget = GetMasterTable(strSection)

Set rngCurrentActivity = tblTarget.ListColumns(2).DataBodyRange.Find(What:=strStartActivity, LookIn:=Excel.XlFindLookIn.xlValues, _
    LookAt:=Excel.XlLookAt.xlWhole)
intTotalDuration = 0

Do While Not rngCurrentActivity Is Nothing
    colActivities.Add rngCurrentActivity.value
    colObjective.Add rngCurrentActivity.Offset(0, 2).value
    colDescription.Add rngCurrentActivity.Offset(0, 3).value
    intActivityDuration = Int(DurationCodeToMins(rngCurrentActivity.Offset(0, 1).value))
    intTotalDuration = intTotalDuration + intActivityDuration
    If intTotalDuration >= intTargetTotalDuration Then
        Set rngNextActivity = rngCurrentActivity
        Set rngCurrentActivity = Nothing
    Else
        If rngCurrentActivity.Offset(1, 0).value = "" Then
            Exit Do
        End If
        Set rngCurrentActivity = rngCurrentActivity.Offset(1, 0)
    End If
Loop

If Not rngNextActivity Is Nothing Then
    strNextActivity = rngNextActivity.value
Else
    strNextActivity = ""
End If

GetNextActivity = strNextActivity

End Function

Function GetBlockDuration(ByVal strBlock As String) As Integer

Dim tblBlocks As ListObject

Set tblBlocks = wsBlocks.ListObjects("tblBlocks")

GetBlockDuration = Application.WorksheetFunction.XLookup(strBlock, tblBlocks.ListColumns("Bloque").DataBodyRange, _
    tblBlocks.ListColumns(4).DataBodyRange, 0)

End Function

Function GetPlanTable(ByVal strSection As String) As ListObject

Dim strGrade As String
Dim strSectionLetter As String
Dim wsTarget As Worksheet
Dim tblTarget As ListObject

If IsNumeric(Left(strSection, 1)) Then
    strGrade = Left(strSection, Len(strSection) - 1)
    strSectionLetter = Right(strSection, 1)
Else
    strGrade = strSection
    strSectionLetter = "A"
End If

Set wsTarget = ThisWorkbook.Worksheets(strGrade & " Actividades TBox")
If strSectionLetter = "A" Then
    Set tblTarget = wsTarget.ListObjects(1)
Else
    Set tblTarget = wsTarget.ListObjects(2)
End If

Set GetPlanTable = tblTarget

End Function

Function GetMasterTable(ByVal strSection As String) As ListObject

Dim strGrade As String
Dim strSectionLetter As String
Dim wsTarget As Worksheet
Dim tblTarget As ListObject

If IsNumeric(Left(strSection, 1)) Then
    strGrade = Left(strSection, Len(strSection) - 1)
    strSectionLetter = Right(strSection, 1)
Else
    strGrade = strSection
    strSectionLetter = "A"
End If

Set wsTarget = wsMasterList
Set tblTarget = wsTarget.ListObjects("tblMasterList" & strGrade)

Set GetMasterTable = tblTarget

End Function


Function GetWorksheetByCodename(wsCodename As String) As Worksheet
    Dim ws As Worksheet
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If ws.CodeName = wsCodename Then
            Set GetWorksheetByCodename = ws
            Exit Function
        End If
    Next ws
    
    ' If no matching worksheet is found, return Nothing
    Set GetWorksheetByCodename = Nothing
End Function
Function GradeNameToNumeric(ByVal strGradeName As String) As String

Select Case LCase(strGradeName)
    Case "primer grado"
        GradeNameToNumeric = "1"
    Case "segundo grado"
        GradeNameToNumeric = "2"
    Case "tercer grado"
        GradeNameToNumeric = "3"
    Case "cuarto grado"
        GradeNameToNumeric = "4"
    Case "quinto grado"
        GradeNameToNumeric = "5"
    Case "sexto grado"
        GradeNameToNumeric = "6"
    Case "s�ptimo grado"
        GradeNameToNumeric = "7"
    Case "octavo grado"
        GradeNameToNumeric = "8"
    Case "noveno grado"
        GradeNameToNumeric = "9"
    Case "d�cimo grado"
        GradeNameToNumeric = "10"
    Case "onceavo grado"
        GradeNameToNumeric = "11"
    Case "doceavo grado"
        GradeNameToNumeric = "12"
    Case "ciclo 3"
        GradeNameToNumeric = "DC3"
End Select

End Function

Function HasDigit(str As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(str)
        If Asc(Mid(str, i, 1)) >= Asc("0") And Asc(Mid(str, i, 1)) <= Asc("9") Then
            HasDigit = True
            Exit Function
        End If
    Next i
    HasDigit = False
End Function
Function ReplaceNumberPlaceholders(ByVal strInput As String, ByVal strGrade As String, _
    ByVal strLessonNumber As String, ByVal strBimesterNumber As String, ByVal strProjectNumber As String) As String

'lesson_number
'bimester_number
'project_number
'project_name

Dim strOutput As String

strOutput = strInput

strOutput = Replace(strOutput, "<lesson_number>", strLessonNumber, 1, -1, vbTextCompare)
strOutput = Replace(strOutput, "<bimester_number>", strBimesterNumber, 1, -1, vbTextCompare)
strOutput = Replace(strOutput, "<project_number>", strProjectNumber, 1, -1, vbTextCompare)

ReplaceNumberPlaceholders = strOutput

End Function

Function ShowDatePicker() As Date
    Dim selectedDate As Date
    
    ' Set the default date to today
    selectedDate = Date
    
    ' Create and show an InputBox with a date format mask
    Dim userInput As String
    Dim strDateFormat As String
    If Application.International(xlMDY) Then
        strDateFormat = "MM/DD/YYYY"
    Else
        strDateFormat = "DD/MM/YYYY"
    End If
    userInput = InputBox("Select a date (" & strDateFormat & "):" & vbNewLine & _
                         "Default is today: " & Format(selectedDate, strDateFormat), _
                         "Date Picker", Format(selectedDate, strDateFormat))
    
    ' Check if user cancelled
    If userInput = "" Then
        ShowDatePicker = CVDate(0) ' Return an empty date
        Exit Function
    End If
    
    ' Try to convert the input to a date
    On Error Resume Next
    selectedDate = CDate(userInput)
    If Err.Number <> 0 Then
        MsgBox "Invalid date format. Please use MM/DD/YYYY.", vbExclamation
        ShowDatePicker = CVDate(0) ' Return an empty date
        Exit Function
    End If
    On Error GoTo 0
    
    ' Return the selected date
    ShowDatePicker = selectedDate
End Function


Function RepeatArrayItems(yInput As Variant, zInput As Variant) As Variant
    Dim yValues As Variant, zValues As Variant
    Dim i As Long, j As Long, k As Long
    Dim outputArray() As Variant
    Dim totalItems As Long
    
    ' Convert inputs to 2D arrays if they're ranges
    If TypeName(yInput) = "Range" Then
        yValues = yInput.value
    Else
        yValues = ConvertTo2DArray(yInput)
    End If
    
    If TypeName(zInput) = "Range" Then
        zValues = zInput.value
    Else
        zValues = ConvertTo2DArray(zInput)
    End If
    
    ' Calculate total items in output
    totalItems = 0
    For i = 1 To UBound(zValues)
        totalItems = totalItems + zValues(i, 1)
    Next i
    
    ' Resize output array
    ReDim outputArray(1 To totalItems, 1 To 1)
    
    ' Fill output array
    k = 1
    For i = 1 To UBound(yValues)
        For j = 1 To zValues(i, 1)
            outputArray(k, 1) = yValues(i, 1)
            k = k + 1
        Next j
    Next i
    
    ' Return the result as a dynamic array
    RepeatArrayItems = outputArray
End Function

Function ConvertTo2DArray(arr As Variant) As Variant
    Dim result() As Variant
    Dim i As Long
    
    If IsArray(arr) Then
        If NumberOfArrayDimensions(arr) = 1 Then
            ReDim result(1 To UBound(arr), 1 To 1)
            For i = 1 To UBound(arr)
                result(i, 1) = arr(i)
            Next i
        Else
            result = arr
        End If
    Else
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = arr
    End If
    
    ConvertTo2DArray = result
End Function

Function NumberOfArrayDimensions(arr As Variant) As Integer
    Dim i As Integer
    Dim tmp As Integer
    
    On Error GoTo FinalDimension
    For i = 1 To 60 ' Arbitrary upper limit
        tmp = UBound(arr, i)
    Next i
FinalDimension:
    NumberOfArrayDimensions = i - 1
End Function
Function AffectsBlocks(ByVal varTimeStart As Variant, varTimeEnd As Variant)

Dim tblBlocks As ListObject
Dim row As ListRow
Dim strBlocksList As String

Set tblBlocks = wsBlocks.ListObjects("tblBlocks")

strBlocksList = ""
For Each row In tblBlocks.ListRows
    If (varTimeStart < row.Range(1, 3).value) And (varTimeEnd > row.Range(1, 2).value) Then
        strBlocksList = strBlocksList & row.Range(1, 1).value & ", "
    End If
Next

If Right(strBlocksList, Len(", ")) = ", " And Len(strBlocksList) > 0 Then
    strBlocksList = Left(strBlocksList, Len(strBlocksList) - Len(", "))
End If

AffectsBlocks = strBlocksList

End Function

Sub GenerateYearlyPlans()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblClassMinutes As ListObject
    Dim wordApp As word.Application
    Dim wordDoc As word.Document
    Dim folderPath As String
    Dim rng As Range
    Dim sec As Range
    Dim sec2 As Range
    Dim sec3 As Range
    Dim sec4 As Range
    Dim templatePath As String
    Dim templateName As String
    Dim newDocName As String
    Dim i As Integer
    Dim j As Integer
    Dim Counter As Integer
    Dim strSection As String
    Dim strBimester As String
    Dim tblClassStatistics As ListObject
    Dim column_offset As Integer
    Dim n_days As Integer
    Dim n_weeks As Integer
    Dim n_hours As Integer
    Dim extracted_text As String
    Dim tblClassesPerSectionPerWeek As ListObject
    Dim number_classes_per_week As Integer
    
    ' Set worksheet and table
    Set ws = wsBimesters
    Set tbl = ws.ListObjects("tblBimesters")
    Set tblClassMinutes = wsClassMinutes.ListObjects("tblClassMinutes")
    Set tblClassStatistics = wsClassStatistics.ListObjects("tblClassStatistics")
    Set tblClassesPerSectionPerWeek = wsSchedule.ListObjects("tblClassesPerSectionPerWeek")
    
    ' Prompt to select or create a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select or create a folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Ask for the Word template path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Word template"
        .Filters.Add "Word Documents", "*.docx"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "No template selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Extract the template name
    templateName = Dir(templatePath)
    
    ' Create a Word application instance
    Set wordApp = New word.Application 'CreateObject(Class:="Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = wdAlertsNone
    
    ' Loop through each section in the table and create a Word document
    For Each sec In tblClassMinutes.ListColumns("Grado").DataBodyRange
        Debug.Print ("Processing section " & sec.value)
        ' Replace "XX" with the actual section in the document name
        newDocName = Replace(templateName, "XX", sec.value)
        
        ' Create a new Word document from the template
        Set wordDoc = wordApp.Documents.Add(templatePath)
        
        With wordDoc
            ' Fill in the fields
            For i = 1 To .ContentControls.count
                Select Case .ContentControls(i).Range.text
                    Case "Section"
                        .ContentControls(i).Range.text = sec.value
                End Select
            Next
            
            Counter = 1
            For Each sec2 In tbl.ListColumns(2).DataBodyRange
                If sec2.value = sec.value Then
                    For j = 1 To .ContentControls.count
                        Select Case .ContentControls(j).Range.text
                            Case "Bimester_" & Trim(str(Counter))
                                strBimester = sec2.Offset(0, -1).value
                                strSection = sec.value
                                .ContentControls(j).Range.text = GetProject(strSection, strBimester)
                            Case "H" & Trim(str(Counter))
                                .ContentControls(j).Range.text = sec2.Offset(0, 1).value
                            Case "B" & Trim(str(Counter))
                                .ContentControls(j).Range.text = CapitalizeFirstLetter(Format(sec2.Offset(0, 2).value, sec2.Offset(0, 2).NumberFormat))
                            Case "E" & Trim(str(Counter))
                                .ContentControls(j).Range.text = CapitalizeFirstLetter(Format(sec2.Offset(0, 3).value, sec2.Offset(0, 3).NumberFormat))
                            Case "BE" & Trim(str(Counter))
                                .ContentControls(j).Range.text = CapitalizeFirstLetter(Format(sec2.Offset(0, 4).value, sec2.Offset(0, 4).NumberFormat))
                        End Select
                    Next
                    Counter = Counter + 1
                End If
                If Counter = 5 Then
                    Exit For
                End If
            Next
            
            For Each sec3 In tblClassStatistics.ListColumns(1).DataBodyRange
                If sec3.value = sec.value Then
                    For j = 1 To .ContentControls.count
                        Select Case Left(.ContentControls(j).Range.text, 1)
                            Case "D"
                                extracted_text = Mid(.ContentControls(j).Range.text, 2, Len(.ContentControls(j).Range.text) - 1)
                                If IsNumeric(extracted_text) Then
                                    column_offset = Val(extracted_text)
                                    n_days = sec3.Offset(0, 1 + column_offset).value
                                    .ContentControls(j).Range.text = n_days
                                End If
                            Case "W"
                                extracted_text = Mid(.ContentControls(j).Range.text, 2, Len(.ContentControls(j).Range.text) - 1)
                                If IsNumeric(extracted_text) Then
                                    column_offset = Val(extracted_text)
                                    n_weeks = sec3.Offset(0, 1 + column_offset).value
                                    .ContentControls(j).Range.text = n_weeks
                                End If
                            Case "H"
                                If Mid(.ContentControls(j).Range.text, 2, 1) = "H" Then
                                    extracted_text = Mid(.ContentControls(j).Range.text, 3, Len(.ContentControls(j).Range.text) - 2)
                                    If IsNumeric(extracted_text) Then
                                        column_offset = Val(extracted_text)
                                        n_hours = sec3.Offset(0, 1 + column_offset).value
                                        .ContentControls(j).Range.text = n_hours
                                    End If
                                End If
                        End Select
                    Next
                End If
            Next
            
            For Each sec4 In tblClassesPerSectionPerWeek.ListColumns(1).DataBodyRange
                If sec4.value = sec.value Then
                    For j = 1 To .ContentControls.count
                        Select Case .ContentControls(j).Range.text
                            Case "NW"
                                number_classes_per_week = sec4.Offset(0, 1).value
                                .ContentControls(j).Range.text = number_classes_per_week
                        End Select
                    Next
                End If
            Next
            
            ' Update the formula field (annotated as 6)
            .Fields.Update
            
            ' Save the document with the new name
            .SaveAs folderPath & "\" & newDocName
            .Close
        End With
        Set wordDoc = Nothing
    Next sec
    
    ' Cleanup
    wordApp.Quit
    Set wordApp = Nothing
    
    MsgBox "Documents created successfully!"
End Sub

Sub GenerateUnitPlans()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblClassMinutes As ListObject
    Dim wordApp As word.Application
    Dim wordDoc As word.Document
    Dim folderPath As String
    Dim rng As Range
    Dim sec As Range
    Dim sec2 As Range
    Dim sec3 As Range
    Dim sec4 As Range
    Dim templatePath As String
    Dim templateName As String
    Dim newDocName As String
    Dim i As Integer
    Dim j As Integer
    Dim Counter As Integer
    Dim strSection As String
    Dim strBimester As String
    Dim tblClassStatistics As ListObject
    Dim column_offset As Integer
    Dim n_days As Integer
    Dim n_weeks As Integer
    Dim n_hours As Integer
    Dim extracted_text As String
    Dim tblClassesPerSectionPerWeek As ListObject
    Dim number_classes_per_week As Integer
    Dim strProjectNumber As String
    Dim strGrade As String
    Dim strBimonthly As String
    
    ' Set worksheet and table
    Set ws = wsBimesters
    Set tbl = ws.ListObjects("tblBimesters")
    Set tblClassMinutes = wsClassMinutes.ListObjects("tblClassMinutes")
    Set tblClassStatistics = wsClassStatistics.ListObjects("tblClassStatistics")
    Set tblClassesPerSectionPerWeek = wsSchedule.ListObjects("tblClassesPerSectionPerWeek")
    
    ' Prompt to select or create a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select or create a folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Ask for the Word template path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Word template"
        .Filters.Add "Word Documents", "*.docx"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "No template selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Extract the template name
    templateName = Dir(templatePath)
    
    ' Create a Word application instance
    Set wordApp = New word.Application 'CreateObject(Class:="Word.Application")
    wordApp.Visible = True
    wordApp.DisplayAlerts = wdAlertsNone
    
    ' Loop through each section in the table and create a Word document
    For Each sec In tblClassMinutes.ListColumns("Grado").DataBodyRange
        Debug.Print ("Processing section " & sec.value)
        
        For Each sec2 In tbl.ListColumns(2).DataBodyRange
            If sec2.value = sec.value Then
                strBimester = sec2.Offset(0, -1).value
                strBimonthly = GetBimonthly(strBimester)
                strSection = sec.value
                strGrade = GetGradeNumberFromSection(strSection)
                strProjectNumber = GetProjectNumberForUnitPlan(strBimester, strSection)
                
                If strProjectNumber <> "" Then
                    ' Create a new Word document from the template
                    Set wordDoc = wordApp.Documents.Add(templatePath)
                    
                    With wordDoc
                        ' Fill placeholders
                        For j = 1 To .ContentControls.count
                            Select Case .ContentControls(j).Range.text
                                Case "YY BIMONTHLY"
                                    .ContentControls(j).Range.text = strBimonthly
                                Case "Date Range"
                                    .ContentControls(j).Range.text = GetDateRange(strBimester, strSection)
                                Case "Section"
                                    .ContentControls(j).Range.text = strSection
                                Case "Modules"
                                    .ContentControls(j).Range.text = GetUnitPlanModules(strProjectNumber, strGrade)
                                Case "Standards"
                                    .ContentControls(j).Range.text = GetUnitPlanStandards(strProjectNumber, strGrade)
                                Case "Contents"
                                    .ContentControls(j).Range.text = GetUnitPlanContents(strProjectNumber, strGrade)
                                Case "Objectives"
                                    .ContentControls(j).Range.text = GetUnitPlanObjectives(strProjectNumber, strGrade)
                                Case "Indicators"
                                    .ContentControls(j).Range.text = GetUnitPlanIndicators(strProjectNumber, strGrade)
                            End Select
                        Next
    
                        ' Update fields and save
                        .Fields.Update
                        newDocName = Replace(templateName, "XX", strSection)
                        newDocName = Replace(newDocName, "BB", strBimonthly)
                        .SaveAs folderPath & "\" & newDocName
                        .Close
                    End With
                    
                    Set wordDoc = Nothing
                End If
            End If
        Next sec2
    Next sec
    
    ' Cleanup
    wordApp.Quit
    Set wordApp = Nothing
    
    MsgBox "Documents created successfully!"
End Sub

Sub GenerateWeeklyPlans(ByVal intWeekNumber As Integer, Optional ByVal varSections As Variant = Empty)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim tblClassMinutes As ListObject
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim folderPath As String
    Dim rng As Range
    Dim sec As Range
    Dim sec2 As Range
    Dim templatePath As String
    Dim templateName As String
    Dim newDocName As String
    Dim i As Integer
    Dim j As Integer
    Dim Counter As Integer
    Dim strSection As String
    Dim strBimester As String
    Dim tblSandbox As ListObject
    Dim strSubject As String
    Dim dicPlanningData As New Dictionary
    Dim tblClassesPerWeek As ListObject
    Dim strBimesterNumber As String
    Dim colTopicListClass1 As New Collection
    Dim colTopicListClass2 As New Collection
    
    strSubject = "Computers"
    
    Set tblClassesPerWeek = wsClassesPerWeek.ListObjects("tblClassesPerWeek")
    strBimester = Application.WorksheetFunction.XLookup(intWeekNumber, tblClassesPerWeek.ListColumns("Semana ABC").DataBodyRange, _
        tblClassesPerWeek.ListColumns("Bimestre").DataBodyRange)
    If strBimester <> "" Then
        strBimesterNumber = Right(strBimester, Len(strBimester) - 1)
    Else
        strBimesterNumber = ""
    End If
    
    ' Set worksheet and table
    Set ws = wsBimesters
    Set tbl = ws.ListObjects("tblBimesters")
    Set tblClassMinutes = wsClassMinutes.ListObjects("tblClassMinutes")
    
    Set tblSandbox = wsSandbox.ListObjects("tblPlanningData")
    
    ' Prompt to select or create a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select or create a folder"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Ask for the Word template path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select the Word template"
        .Filters.Add "Word Documents", "*.docx"
        If .Show = -1 Then
            templatePath = .SelectedItems(1)
        Else
            MsgBox "No template selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Extract the template name
    templateName = Dir(templatePath)
    
    ' Create a Word application instance
    On Error Resume Next
    Set wordApp = GetObject(Class:="Word.Application")
    If wordApp Is Nothing Then
        Set wordApp = CreateObject(Class:="Word.Application")
    End If
    On Error GoTo 0
    
    DeleteAllRowsInTable tblSandbox
    
    ' Loop through each section in the table and create a Word document
    For Each sec In tblClassMinutes.ListColumns("Grado").DataBodyRange
        strSection = sec
        If sec <> "" Then
            If Not IsSectionIncluded(CStr(sec.value), varSections) Then
                GoTo NextItem
            End If
            Debug.Print ("Processing section " & strSection)
            
            FillPlanningDataRecord intWeekNumber, strSection
            Set dicPlanningData = ReadPlanningRecord(intWeekNumber, strSection)
            
            ' Replace "XX" with the actual section in the document name
            newDocName = Replace(templateName, "XX", sec.value)
            
            ' Create a new Word document from the template
            Set wordDoc = wordApp.Documents.Add(templatePath)
            
            With wordDoc
                ' Fill in the fields
                For i = .ContentControls.count To 1 Step -1
                    Select Case .ContentControls(i).Range.text
                        Case "<section>"
                            .ContentControls(i).Range.text = sec.value
                        Case "<subject>"
                            .ContentControls(i).Range.text = strSubject
                        Case "<week_number>"
                            .ContentControls(i).Range.text = intWeekNumber
                        Case "<bimester_number>"
                            .ContentControls(i).Range.text = strBimesterNumber
                        Case "<start_date>"
                            .ContentControls(i).Range.text = FormatDateCustom(dicPlanningData.Item("<start_date>"))
                        Case "<end_date>"
                            .ContentControls(i).Range.text = FormatDateCustom(dicPlanningData.Item("<end_date>"))
                        Case "<topic_class_1>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<topic_class_1>")
                        Case "<topic_class_2>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<topic_class_2>")
                        Case "<objective_class_1>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<objective_class_1>")
                        Case "<objective_class_2>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<objective_class_2>")
                        Case "<project_name_class_1>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<project_name_class_1>")
                        Case "<project_name_class_2>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<project_name_class_2>")
                        Case "<activity_number_class_1>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<activity_number_class_1>")
                        Case "<activity_number_class_2>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<activity_number_class_2>")
                        Case "<description_class_1>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<description_class_1>")
                        Case "<description_class_2>"
                            .ContentControls(i).Range.text = dicPlanningData.Item("<description_class_2>")
                        Case "<topic_list_class_1>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<topic_list_class_1>")
                        Case "<topic_list_class_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<topic_list_class_2>")
                        Case "<topic_list_class_1_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<topic_list_class_1_2>")
                        Case "<objective_list_class_1>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<objective_list_class_1>")
                        Case "<objective_list_class_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<objective_list_class_2>")
                        Case "<objective_list_class_1_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<objective_list_class_1_2>")
                        Case "<description_list_class_1>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<description_list_class_1>")
                        Case "<description_list_class_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<description_list_class_2>")
                        Case "<description_list_class_1_2>"
                            ReplacePlaceholderWithNumberedListInRichText .ContentControls(i), dicPlanningData.Item("<description_list_class_1_2>")
                    End Select
                Next
                
                ' Update the formula field (annotated as 6)
                .Fields.Update
                
                ' Save the document with the new name
                .SaveAs folderPath & "\" & newDocName
                .Close
            End With
        End If
NextItem:
    Next sec
    
PrematureExit:
    
    ' Cleanup
    On Error Resume Next
    wordApp.Quit
    On Error GoTo 0
    Set wordApp = Nothing
    
    MsgBox "Documents created successfully!"
End Sub

Sub ReplacePlaceholderWithNumberedListInCollection(ByVal doc As word.Document, ByVal cc As word.ContentControl, ByVal colActivities As Collection)
    Dim i As Integer
    Dim activity As Variant
    Dim paragraphRange As word.Range
    
    ' Clear the existing content in the rich text content control
    cc.Range.text = ""
    
    ' Loop through the collection and insert each activity as a numbered item
    For i = 1 To colActivities.count
        ' Add a new paragraph within the content control range
        Set paragraphRange = cc.Range.Paragraphs.Add().Range
        paragraphRange.text = colActivities(i)
        
        ' Apply default numbered list formatting
        paragraphRange.ListFormat.ApplyNumberDefault
    Next i
    
End Sub

Sub ReplacePlaceholderWithNumberedListInRichText(ByVal cc As word.ContentControl, ByVal strActivities As String)
    Dim i As Integer
    Dim arrActivities() As String
    Dim paragraphRange As word.Range
    
   ' Clear the existing content in the rich text content control
    cc.Range.text = ""
    
    ' Add a new paragraph within the content control range
    Set paragraphRange = cc.Range.Paragraphs.Add().Range
    paragraphRange.text = Trim(strActivities)  ' Remove any extra spaces
    
    ' Apply default numbered list formatting
    paragraphRange.ListFormat.ApplyNumberDefault

End Sub


Sub DeleteAllRowsInTable(tbl As ListObject)

    ' Remove any filters
    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData

    ' Clear the contents of the DataBodyRange (if it exists)
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
        tbl.DataBodyRange.Delete
    End If
End Sub
Function ReadPlanningRecord(ByVal intWeekNumber As Integer, ByVal strSection As String) As Scripting.Dictionary

Dim rngFoundRecord As Range
Dim tblPlanningData As ListObject
Dim intLastRowIndex As Integer
Dim dicData As New Scripting.Dictionary
Dim datStartDate As Date
Dim datEndDate As Date
Dim strTopicClass1 As String
Dim strTopicClass2 As String
Dim strObjectiveClass1 As String
Dim strObjectiveClass2 As String
Dim intProjectNumberClass1 As Integer
Dim intProjectNumberClass2 As Integer
Dim strProjectNameClass1 As String
Dim strProjectNameClass2 As String
Dim intActivityNumberClass1 As Integer
Dim intActivityNumberClass2 As Integer
Dim strActivityNameClass1 As String
Dim strActivityNameClass2 As String
Dim strDescriptionClass1 As String
Dim strDescriptionClass2 As String
Dim strTopicListClass1 As String
Dim strTopicListClass2 As String

Set tblPlanningData = wsSandbox.ListObjects("tblPlanningData")

tblPlanningData.Range.AutoFilter 1, intWeekNumber
tblPlanningData.Range.AutoFilter 5, strSection

Set rngFoundRecord = tblPlanningData.DataBodyRange.SpecialCells(xlCellTypeVisible)
intLastRowIndex = rngFoundRecord.Rows.count

If intLastRowIndex > 0 Then
    dicData.Add "<start_date>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<start_date>")).value
    dicData.Add "<end_date>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<end_date>")).value
    dicData.Add "<topic_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<topic_class_1>")).value
    dicData.Add "<objective_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<objective_class_1>")).value
    dicData.Add "<project_tivitynumber_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<project_tivitynumber_class_1>")).value
    dicData.Add "<project_name_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<project_name_class_1>")).value
    dicData.Add "<ac_number_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<ac_number_class_1>")).value
    dicData.Add "<activity_name_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<activity_name_class_1>")).value
    dicData.Add "<description_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<description_class_1>")).value
    dicData.Add "<topic_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<topic_class_2>")).value
    dicData.Add "<objective_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<objective_class_2>")).value
    dicData.Add "<project_number_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<project_number_class_2>")).value
    dicData.Add "<project_name_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<project_name_class_2>")).value
    dicData.Add "<activity_number_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<activity_number_class_2>")).value
    dicData.Add "<activity_name_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<activity_name_class_2>")).value
    dicData.Add "<description_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<description_class_2>")).value
    dicData.Add "<topic_list_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<topic_list_class_1>")).value
    dicData.Add "<topic_list_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<topic_list_class_2>")).value
    dicData.Add "<topic_list_class_1_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<topic_list_class_1_2>")).value
    dicData.Add "<objective_list_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<objective_list_class_1>")).value
    dicData.Add "<objective_list_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<objective_list_class_2>")).value
    dicData.Add "<objective_list_class_1_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<objective_list_class_1_2>")).value
    dicData.Add "<description_list_class_1>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<description_list_class_1>")).value
    dicData.Add "<description_list_class_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<description_list_class_2>")).value
    dicData.Add "<description_list_class_1_2>", rngFoundRecord(1, GetColumnNumber(tblPlanningData, "<description_list_class_1_2>")).value
Else
    dicData.Add "<start_date>", ""
    dicData.Add "<end_date>", ""
    dicData.Add "<topic_class_1>", ""
    dicData.Add "<objective_class_1>", ""
    dicData.Add "<project_tivitynumber_class_1>", ""
    dicData.Add "<project_name_class_1>", ""
    dicData.Add "<ac_number_class_1>", ""
    dicData.Add "<activity_name_class_1>", ""
    dicData.Add "<description_class_1>", ""
    dicData.Add "<topic_class_2>", ""
    dicData.Add "<objective_class_2>", ""
    dicData.Add "<project_number_class_2>", ""
    dicData.Add "<project_name_class_2>", ""
    dicData.Add "<activity_number_class_2>", ""
    dicData.Add "<activity_name_class_2>", ""
    dicData.Add "<description_class_2>", ""
    dicData.Add "<topic_list_class_1>", ""
    dicData.Add "<topic_list_class_2>", ""
    dicData.Add "<topic_list_class_1_2>", ""
    dicData.Add "<objective_list_class_1>", ""
    dicData.Add "<objective_list_class_2>", ""
    dicData.Add "<objective_list_class_1_2>", ""
    dicData.Add "<description_list_class_1>", ""
    dicData.Add "<description_list_class_2>", ""
    dicData.Add "<description_list_class_1_2>", ""
End If
Set ReadPlanningRecord = dicData

End Function

Function FormatDateCustom(inputDate As Date) As String
    Dim dayNumber As Integer
    Dim daySuffix As String
    Dim monthName As String
    
    ' Get the day number and the three-letter month name
    dayNumber = Day(inputDate)
    monthName = Format(inputDate, "mmm")
    
    ' Determine the suffix (st, nd, rd, th)
    Select Case dayNumber
        Case 1, 21, 31
            daySuffix = "st"
        Case 2, 22
            daySuffix = "nd"
        Case 3, 23
            daySuffix = "rd"
        Case Else
            daySuffix = "th"
    End Select
    
    ' Return the formatted string
    FormatDateCustom = monthName & " " & dayNumber & daySuffix
End Function


Function CapitalizeFirstLetter(ByVal str As String) As String
    If Len(str) > 0 Then
        CapitalizeFirstLetter = UCase(Left(str, 1)) & LCase(Mid(str, 2))
    Else
        CapitalizeFirstLetter = ""
    End If
End Function
Function GetDateCategory(ByVal rngDate As Range)

Dim tblClassInterruptions As ListObject

Set tblClassInterruptions = wsClassInterruptions.ListObjects("tblClassInterruptions")



End Function

Function GetProject(grade As String, bimester As String) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim projectRange As ListRows
    Dim row As ListRow
    Dim found As Boolean
    found = False

    ' Set the worksheet based on the sheet name in VBA properties
    Set ws = wsProjects

    ' Set the table object based on the table name
    Set tbl = ws.ListObjects("tblProjects")
    Set projectRange = tbl.ListRows

    ' Loop through each row in the table to find the matching grade and bimester
    For Each row In projectRange
        If row.Range(1, 1).value = grade And row.Range(1, 2).value = bimester Then
            GetProject = row.Range(1, 3).value
            found = True
            Exit Function
        End If
    Next row
    
    ' If no match is found
    If Not found Then
        GetProject = "Project not found"
    End If
End Function

Function GetSectionsAndBlocksAffected(ByVal rngDateStart As Range, rngDateEnd As Range, _
    varTimeStart As Variant, varTimeEnd As Variant, rngReason As Range)
    
Dim booIsSenior As Boolean
Dim answer As String
Dim strBlocksAffected As String

If InStr(1, rngReason.value, "Seniors", vbTextCompare) Then
    booIsSenior = True
End If

If booIsSenior = False And varTimeStart = "N/A" And varTimeEnd = "N/A" Then
    answer = "Todos"
ElseIf booIsSenior Then
    answer = "12A"
Else
    strBlocksAffected = AffectsBlocks(varTimeStart, varTimeEnd)
    If strBlocksAffected <> "" Then
        answer = strBlocksAffected
    End If
End If

GetSectionsAndBlocksAffected = answer
    
End Function

Sub UpdateClassMinutes()

Dim tblClassMinutes As ListObject
Dim tblBlocks As ListObject
Dim tblSchedule As ListObject
Dim rngGrade As Range
Dim rngFoundCell As Range
Dim rngSearchRange As Range
Dim rngFoundGrade As Range
Dim strFirstAddress As String
Dim strClassMinutes As String
Dim intClassMinutes As Integer
Dim strDay As String
Dim varBlock As Variant
Dim intMinutes As Integer

Set tblClassMinutes = wsClassMinutes.ListObjects("tblClassMinutes")
Set tblBlocks = wsBlocks.ListObjects("tblBlocks")
Set tblSchedule = wsSchedule.ListObjects("tblSchedule")

Set rngSearchRange = tblSchedule.DataBodyRange

For Each rngGrade In tblClassMinutes.ListColumns(1).DataBodyRange
    strClassMinutes = ""
    Set rngFoundCell = rngSearchRange.Find(What:=rngGrade.value, LookIn:=Excel.xlValues, _
        LookAt:=Excel.xlWhole)
    If Not rngFoundCell Is Nothing Then
        strFirstAddress = rngFoundCell.Address
        intClassMinutes = 0
        Do
            strDay = tblSchedule.HeaderRowRange(rngFoundCell.column - tblSchedule.HeaderRowRange(1, 1).column + 1)
            varBlock = tblSchedule.ListColumns(1).DataBodyRange(rngFoundCell.row - tblSchedule.HeaderRowRange(1, 1).row).value
            strClassMinutes = strClassMinutes & ", " & "(" & strDay & " Block " & varBlock & ")"
            intMinutes = WorksheetFunction.XLookup(varBlock, tblBlocks.ListColumns(1).DataBodyRange, tblBlocks.ListColumns(4).DataBodyRange)
            intClassMinutes = intClassMinutes + intMinutes
            Set rngFoundCell = rngSearchRange.FindNext(rngFoundCell)
        Loop While Not rngFoundCell Is Nothing And rngFoundCell.Address <> strFirstAddress
        If Left(strClassMinutes, 2) = ", " Then
            strClassMinutes = Right(strClassMinutes, Len(strClassMinutes) - 2)
        End If
        Set rngFoundGrade = tblClassMinutes.ListColumns(1).Range.Find(What:=rngGrade.value, LookIn:=Excel.xlValues, LookAt:=Excel.xlWhole)
        rngFoundGrade.Offset(0, 1).value = strClassMinutes
        ' rngFoundGrade.Offset(0, 2).Value = intClassMinutes No longer needed, replaced by formulas.
        'Debug.Print rngFoundCell.Value, strClassMinutes, intClassMinutes
    Else
        Debug.Print "No instances found for " & rngGrade.value
    End If
Next

End Sub


Sub FillPlanningDataRecord(ByVal intWeekNumber As Integer, ByVal strSection As String)

Dim tblPlanningData As ListObject
Dim tblClassesPerWeek As ListObject
Dim rowRecord As ListRow
Dim datStartDate As Date
Dim datEndDate As Date
Dim strBlock1 As Variant
Dim strBlock2 As Variant
Dim strActivityName1 As String
Dim strActivityName2 As String
Dim strProjectNumber1 As String
Dim strActivityNumber1 As String
Dim strProjectNumber2 As String
Dim strActivityNumber2 As String
Dim strObjective1 As String
Dim strDescription1 As String
Dim strObjective2 As String
Dim strDescription2 As String
Dim r As Variant
Dim strGrade As String
Dim strProjectName1 As String
Dim strProjectName2 As String
Dim strBimester As String
Dim intBlock1Duraction As Integer
Dim intBlock2Duration As Integer
Dim colActivityListClass1 As New Collection
Dim colActivityListClass2 As New Collection
Dim strTopicListClass1 As String
Dim strTopicListClass2 As String
Dim strTopicListClass1_2 As String
Dim strDescriptionListClass1 As String
Dim strDescriptionListClass2 As String
Dim strDescriptionListClass1_2 As String
Dim rActivityData1 As Variant
Dim rActivityData2 As Variant
Dim colObjectiveListClass1 As New Collection
Dim colObjectiveListClass2 As New Collection
Dim colDescriptionListClass1 As New Collection
Dim colDescriptionListClass2 As New Collection
Dim strObjectiveListClass1 As String
Dim strObjectiveListClass2 As String
Dim strObjectiveListClass1_2 As String
Dim booIsPastWeek As Boolean

Set tblPlanningData = wsSandbox.ListObjects("tblPlanningData")
Set tblClassesPerWeek = wsClassesPerWeek.ListObjects("tblClassesPerWeek")
strBimester = Application.WorksheetFunction.XLookup(intWeekNumber, tblClassesPerWeek.ListColumns("Semana ABC").DataBodyRange, _
    tblClassesPerWeek.ListColumns("Bimestre").DataBodyRange)

Set rowRecord = tblPlanningData.ListRows.Add(position:=tblPlanningData.ListRows.count + 1, AlwaysInsert:=True)

rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<week_number>")) = intWeekNumber
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<grade>")) = strSection
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<bimester_number>")) = strBimester
datStartDate = Application.WorksheetFunction.XLookup(intWeekNumber, tblClassesPerWeek.ListColumns("Semana ABC").Range, tblClassesPerWeek.ListColumns("Fecha Inicio").Range)
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<start_date>")) = datStartDate
datEndDate = Application.WorksheetFunction.XLookup(intWeekNumber, tblClassesPerWeek.ListColumns("Semana ABC").Range, tblClassesPerWeek.ListColumns("Fecha Fin").Range)
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<end_date>")) = datEndDate
strBlock1 = GetBlockNumber(intWeekNumber, strSection, 1)
strBlock2 = GetBlockNumber(intWeekNumber, strSection, 2)
intBlock1Duraction = GetBlockDuration(strBlock1)
intBlock2Duration = GetBlockDuration(strBlock2)
strGrade = Left(strSection, Len(strSection) - 1)

If strBlock1 = "" Then
    strActivityName1 = ""
    strProjectNumber1 = ""
    strActivityNumber1 = ""
    strObjective1 = ""
    strDescription1 = ""
    strProjectName1 = ""
Else
    strActivityName1 = GetActivityOnPastWeek(datStartDate, datEndDate, strSection, strBlock1)
    If strActivityName1 = "" Then
        strActivityName1 = GetActivityOnFutureWeek(datStartDate, datEndDate, strSection, strBlock1, intWeekNumber)
    End If
    r = GetProjectAndActivityNumber(strActivityName1)
    strProjectNumber1 = r(0)
    strActivityNumber1 = r(1)
    r = GetTBoxProjectInfo(strProjectNumber1, strActivityNumber1, strGrade)
    strObjective1 = r(0)
    strDescription1 = r(1)
    strProjectName1 = r(2)
    rActivityData1 = GetActivityList(CleanActivityName(strActivityName1), strSection, strBlock1)
    Set colActivityListClass1 = rActivityData1(0)
    Set colObjectiveListClass1 = rActivityData1(1)
    Set colDescriptionListClass1 = rActivityData1(2)
    strTopicListClass1 = Replace(CollToLineList(colActivityListClass1), "<project_name>", strProjectName1, 1, -1, vbTextCompare)
    strObjectiveListClass1 = Replace(CollToLineList(colObjectiveListClass1), "<project_name>", strProjectName1, 1, -1, vbTextCompare)
    strDescriptionListClass1 = Replace(CollToLineList(colDescriptionListClass1), "<project_name>", strProjectName1, 1, -1, vbTextCompare)
End If
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<topic_class_1>")) = strActivityName1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<activity_name_class_1>")) = strActivityName1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<project_number_class_1>")) = strProjectNumber1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<activity_number_class_1>")) = strActivityNumber1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<objective_class_1>")) = strObjective1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<description_class_1>")) = strDescription1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<project_name_class_1>")) = strProjectName1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<topic_list_class_1>")) = strTopicListClass1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<objective_list_class_1>")) = strObjectiveListClass1
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<description_list_class_1>")) = strDescriptionListClass1

If strBlock2 = "" Then
    strActivityName2 = ""
    strProjectNumber2 = ""
    strActivityNumber2 = ""
    strObjective2 = ""
    strDescription2 = ""
    strProjectName2 = ""
Else
    strActivityName2 = GetActivityOnPastWeek(datStartDate, datEndDate, strSection, strBlock2)
    If strActivityName2 = "" Then
        strActivityName2 = GetActivityOnFutureWeek(datStartDate, datEndDate, strSection, strBlock2, intWeekNumber)
    End If
    r = GetProjectAndActivityNumber(strActivityName2)
    strProjectNumber2 = r(0)
    strActivityNumber2 = r(1)
    r = GetTBoxProjectInfo(strProjectNumber1, strActivityNumber2, strGrade)
    strObjective2 = r(0)
    strDescription2 = r(1)
    strProjectName2 = r(2)
    rActivityData2 = GetActivityList(CleanActivityName(strActivityName1), strSection, strBlock1)
    Set colActivityListClass2 = rActivityData2(0)
    Set colObjectiveListClass2 = rActivityData2(1)
    Set colDescriptionListClass2 = rActivityData2(2)
    strTopicListClass2 = Replace(CollToLineList(colActivityListClass2), "<project_name>", strProjectName2, 1, -1, vbTextCompare)
    strObjectiveListClass2 = Replace(CollToLineList(colObjectiveListClass2), "<project_name>", strProjectName2, 1, -1, vbTextCompare)
    strDescriptionListClass2 = Replace(CollToLineList(colDescriptionListClass2), "<project_name>", strProjectName2, 1, -1, vbTextCompare)
End If

rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<topic_class_2>")) = strActivityName2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<activity_name_class_2>")) = strActivityName2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<project_number_class_2>")) = strProjectNumber2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<activity_number_class_2>")) = strActivityNumber2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<objective_class_2>")) = strObjective2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<description_class_2>")) = strDescription2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<project_name_class_2>")) = strProjectName2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<topic_list_class_2>")) = strTopicListClass2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<objective_list_class_2>")) = strObjectiveListClass2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<description_list_class_2>")) = strDescriptionListClass2

If strTopicListClass1 = strTopicListClass2 Then
    strTopicListClass1_2 = strTopicListClass1
Else
    strTopicListClass1_2 = strTopicListClass1 & vbCrLf & strTopicListClass2
End If

If strObjectiveListClass1 = strObjectiveListClass2 Then
    strObjectiveListClass1_2 = strObjectiveListClass1
Else
    strObjectiveListClass1_2 = strObjectiveListClass1 & vbCrLf & strObjectiveListClass2
End If

If strDescriptionListClass1 = strDescriptionListClass2 Then
    strDescriptionListClass1_2 = strDescriptionListClass1
Else
    strDescriptionListClass1_2 = strDescriptionListClass1 & vbCrLf & strDescriptionListClass2
End If

rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<topic_list_class_1_2>")) = strTopicListClass1_2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<objective_list_class_1_2>")) = strObjectiveListClass1_2
rowRecord.Range(1, GetColumnNumber(tblPlanningData, "<description_list_class_1_2>")) = strDescriptionListClass1_2

End Sub

Function GetNumberedList(colActivities As Collection) As String
    Dim numberedList As String
    Dim i As Integer
    Dim activity As Variant
    
    ' Initialize the string to store the numbered list
    numberedList = ""
    
    ' Loop through the collection and create a numbered list
    For i = 1 To colActivities.count
        ' Add the numbered item followed by a line break (vbCrLf)
        numberedList = numberedList & i & ". " & colActivities(i) & vbCrLf
    Next i
    
    ' Remove the last line break for cleaner output
    If Len(numberedList) > 0 Then
        numberedList = Left(numberedList, Len(numberedList) - Len(vbCrLf))
    End If
    
    ' Return the formatted numbered list
    GetNumberedList = numberedList
End Function

Function CollToLineList(colActivities As Collection) As String
    Dim strLineList As String
    Dim i As Integer
    Dim activity As Variant
    
    ' Initialize the string to store the line list
    strLineList = ""
    
    ' Loop through the collection and create a line list
    For i = 1 To colActivities.count
        ' Add the numbered item followed by a line break (vbCrLf)
        strLineList = strLineList & colActivities(i) & vbCrLf
    Next i
    
    ' Remove the last line break for cleaner output
    If Len(strLineList) > 0 Then
        strLineList = Left(strLineList, Len(strLineList) - Len(vbCrLf))
    End If
    
    ' Return the formatted line list
    CollToLineList = strLineList
End Function
Function CollectionToLineList(ByVal col As Collection, ByVal strSeparator As String)

Dim strOutput As String
Dim i As Integer

For i = 1 To col.count
    strOutput = strOutput & col(i) & strSeparator
Next
If Right(strOutput, Len(strSeparator)) = strSeparator Then
    strOutput = Left(strOutput, Len(strOutput) - Len(strSeparator))
End If

CollectionToLineList = strOutput

End Function

Function LineListToCollection(ByVal strInput As String, ByVal strSeparator As String) As Collection
    Dim colOutput As New Collection
    Dim arrItems() As String
    Dim i As Integer

    ' Split the input string into an array using the specified separator
    arrItems = Split(strInput, strSeparator)

    ' Loop through the array and add each item to the collection
    For i = LBound(arrItems) To UBound(arrItems)
        colOutput.Add Trim(arrItems(i))  ' Trim to remove any extra spaces
    Next i

    ' Return the populated collection
    Set LineListToCollection = colOutput
End Function


Function CleanActivityName(ByVal strFullActivityName As String) As String

Dim intColonPos As Integer
Dim strCleandedActivityName As String

intColonPos = InStr(1, strFullActivityName, ":", vbTextCompare)
If intColonPos <> 0 Then
    strCleandedActivityName = Left(strFullActivityName, intColonPos - 1)
Else
    strCleandedActivityName = strFullActivityName
End If

CleanActivityName = strCleandedActivityName

End Function

Function GetTBoxProjectInfo(ByVal strProjectNumber As String, ByVal strActivityNumber As String, ByVal strGrade As String) As Variant

Dim tblTBoxActivities As ListObject
Dim strObjective As String
Dim strDescription As String
Dim r As Variant
Dim rngFilteredRange As Range
Dim strProjectName As String

Set tblTBoxActivities = wsTBoxActivities.ListObjects("tblTBoxActivities")

If strProjectNumber <> "" And strActivityNumber <> "" Then

    tblTBoxActivities.Range.AutoFilter Field:=1, Criteria1:=strGrade
    tblTBoxActivities.Range.AutoFilter Field:=2, Criteria1:=strProjectNumber
    tblTBoxActivities.Range.AutoFilter Field:=5, Criteria1:=strActivityNumber
    On Error Resume Next
    Set rngFilteredRange = tblTBoxActivities.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not rngFilteredRange Is Nothing Then
        strObjective = rngFilteredRange(1, 7)
        strDescription = rngFilteredRange(1, 8)
        strProjectName = rngFilteredRange(1, 3)
    Else
        strObjective = ""
        strDescription = ""
        strProjectName = ""
    End If

Else

    strObjective = ""
    strDescription = ""
    strProjectName = ""
End If
    

GetTBoxProjectInfo = Array(strObjective, strDescription, strProjectName)

End Function

Function ApplyTwoLevelAutoFilterToArray(dataRange As Range, col1Index As Long, col1Criteria As String, _
                                        col2Index As Long, col2Criteria As String, returnColIndex As Long) As Variant
    Dim filteredRange As Range
    Dim cell As Range
    Dim result() As Variant
    Dim rowCount As Long
    Dim i As Long
    
    ' Check if data range has an AutoFilter
    If Not dataRange.Parent.AutoFilterMode Then
        dataRange.AutoFilter
    End If
    
    ' Apply the first filter condition (Column 1)
    dataRange.AutoFilter Field:=col1Index, Criteria1:=col1Criteria
    
    ' Apply the second filter condition (Column 2)
    dataRange.AutoFilter Field:=col2Index, Criteria1:=col2Criteria
    
    ' Set the filtered range (excluding headers)
    On Error Resume Next
    Set filteredRange = dataRange.SpecialCells(xlCellTypeVisible).Resize(dataRange.Rows.count - 1).Offset(1, 0)
    On Error GoTo 0
    
    ' Check if there are visible cells
    If filteredRange Is Nothing Then
        ' Return an empty array if no rows are visible
        ApplyTwoLevelAutoFilterToArray = Array()
        Exit Function
    End If
    
    ' Get the number of rows in the filtered data
    rowCount = filteredRange.Rows.count
    
    ' Resize the result array to hold the values from the specified column
    ReDim result(1 To rowCount)
    
    ' Populate the result array with the visible data from the specified return column
    i = 1 ' Row counter for the result array
    For Each cell In filteredRange.Areas(1).Rows
        result(i) = cell.Cells(1, returnColIndex).value ' Get the value from the return column
        i = i + 1
    Next cell
    
    ' Return the array of filtered data
    ApplyTwoLevelAutoFilterToArray = result
End Function





Function GetProjectAndActivityNumber(inputText As String) As Variant
    Dim projectNumber As String
    Dim activityNumber As String
    Dim result As Variant
    Dim parts As Variant
    Dim delimiterPos As Long
    Dim cleanText As String
    Dim projectPart As String
    Dim activityPart As String
    Dim intColonPos As Integer
    
    intColonPos = InStr(1, inputText, ":", vbTextCompare)
    If intColonPos <> 0 Then
        cleanText = Left(inputText, intColonPos - 1)
    Else
        cleanText = inputText
    End If
    
    parts = Split(cleanText, ", ")
    If UBound(parts) = 1 Then
        projectPart = parts(0)
        activityPart = parts(1)
        parts = Split(projectPart, " ")
        projectNumber = parts(1)
        parts = Split(activityPart, " ")
        activityNumber = parts(1)
    Else
        If UBound(parts) = 0 Then
            If Left(parts(0), Len("Project ")) = "Project " Then
                parts = Split(parts(0), " ")
                projectNumber = parts(1)
                activityNumber = ""
            Else
                projectNumber = ""
                activityNumber = ""
            End If
        Else
            projectNumber = ""
            activityNumber = ""
        End If
    End If
    
    result = Array(projectNumber, activityNumber)
    
    GetProjectAndActivityNumber = result
    
End Function

Function GetActivityOnPastWeek(ByVal datWeekStartDate As Date, ByVal datWeekEndDate As Date, ByVal strSection As String, ByVal strBlock As String)

Dim wsTarget As Worksheet
Dim strGrade As String
Dim strSectionLetter As String
Dim tblTarget As ListObject
Dim rngDate As Range
Dim rngBlockItem As Range
Dim intRowIndex As Integer
Dim strActivity As String

Set tblTarget = GetPlanTable(strSection)

For Each rngDate In tblTarget.ListColumns("Date").DataBodyRange
    If (rngDate.value >= datWeekStartDate) And (rngDate.value <= datWeekEndDate) Then
        intRowIndex = rngDate.row - tblTarget.HeaderRowRange.row
        Set rngBlockItem = tblTarget.ListColumns("Block").DataBodyRange(intRowIndex, 1)
        If rngBlockItem.value = strBlock Then
            strActivity = tblTarget.ListColumns("Current Activity").DataBodyRange(intRowIndex, 1)
        End If
    End If
Next

GetActivityOnPastWeek = strActivity

End Function

Function GetActivityOnFutureWeek(ByVal datWeekStartDate As Date, ByVal datWeekEndDate As Date, ByVal strSection As String, ByVal strBlock As String, _
    intTargetWeek As Integer)

Dim wsTarget As Worksheet
Dim strGrade As String
Dim strSectionLetter As String
Dim tblTarget As ListObject
Dim rngDate As Range
Dim rngBlockItem As Range
Dim intRowIndex As Integer
Dim strActivity As String
Dim strLastActivity As String
Dim r As Variant
Dim strLastActivityBlock As String
Dim datLastActivityDate As Date
Dim rngBlocksInBetween As Range
Dim i As Integer
Dim intItemBlockDuration As Integer
Dim strItemBlock As String
Dim intTotalDuration As Integer

Set tblTarget = GetPlanTable(strSection)
r = GetLastActivityDone(tblTarget)
strLastActivity = r(0)
strLastActivityBlock = r(1)
datLastActivityDate = r(2)
'Debug.Print strLastActivity, strLastActivityBlock, datLastActivityDate

wsHelper.Names("Section").RefersToRange.value = strSection
wsHelper.Names("Last_Activity_Date").RefersToRange.value = datLastActivityDate
wsHelper.Names("Target_Week").RefersToRange.value = intTargetWeek
Set rngBlocksInBetween = wsHelper.Range("Block_List")

intTotalDuration = 0
For i = 1 To rngBlocksInBetween.Columns.count
    strItemBlock = rngBlocksInBetween(1, i).value
    intItemBlockDuration = GetBlockDuration(strItemBlock)
    'Debug.Print strItemBlock, intItemBlockDuration
    intTotalDuration = intTotalDuration + intItemBlockDuration
Next
'Debug.Print intTotalDuration

GetActivityOnFutureWeek = GetNextActivity(strLastActivity, intTotalDuration, strSection)

End Function



Function GetLastActivityDone(ByVal tblTarget As ListObject) As Variant

Dim i As Integer
Dim strLastActivity As String
Dim strBlock As String
Dim datDate As Date

strLastActivity = ""
For i = tblTarget.DataBodyRange.Rows.count To 1 Step -1
    If tblTarget.ListColumns("Completed").DataBodyRange(i).value = 1 Then
        strLastActivity = tblTarget.ListColumns("Current Activity").DataBodyRange(i).Offset(-1, 1).value
        strBlock = tblTarget.ListColumns("Block").DataBodyRange(i).value
        datDate = tblTarget.ListColumns("Date").DataBodyRange(i).value
        Exit For
    End If
Next

GetLastActivityDone = Array(strLastActivity, strBlock, datDate)

End Function

Function GetBlockNumber(ByVal intWeekNumber As Integer, ByVal strSection As String, ByVal intSessionNumber As Integer) As String

Dim tblClassesPerWeek As ListObject
Dim r As Variant

Set tblClassesPerWeek = wsClassesPerWeek.ListObjects("tblClassesPerWeek")

r = Application.WorksheetFunction.Filter(tblClassesPerWeek.ListColumns(strSection).DataBodyRange, _
    wsClassesPerWeek.Evaluate(tblClassesPerWeek.ListColumns("Semana ABC").DataBodyRange.Address & "=" & str(intWeekNumber)))
    
If UBound(r, 1) > intSessionNumber - 1 Then
    GetBlockNumber = r(intSessionNumber, 1)
Else
    GetBlockNumber = ""
End If

End Function

Private Function IsSectionIncluded(ByVal strSection As String, ByVal varSections As Variant) As Boolean
    ' Returns True if strSection should be processed: when varSections is Empty, include all;
    ' otherwise include only when strSection is in the list (single string or array), case-insensitive.
    If IsEmpty(varSections) Then
        IsSectionIncluded = True
        Exit Function
    End If
    Dim i As Long
    If IsArray(varSections) Then
        For i = LBound(varSections) To UBound(varSections)
            If StrComp(CStr(varSections(i)), strSection, vbTextCompare) = 0 Then
                IsSectionIncluded = True
                Exit Function
            End If
        Next i
    Else
        If StrComp(CStr(varSections), strSection, vbTextCompare) = 0 Then
            IsSectionIncluded = True
            Exit Function
        End If
    End If
    IsSectionIncluded = False
End Function

Function GetColumnNumber(ByVal tbl As ListObject, strColumnName As String)

Dim rngFound As Range
Dim intColNumber As Integer

Set rngFound = tbl.HeaderRowRange.Find(What:=strColumnName, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
If Not rngFound Is Nothing Then
    intColNumber = rngFound.column - tbl.HeaderRowRange(1, 1).column + 1
Else
    intColNumber = 0
End If

GetColumnNumber = intColNumber

End Function

Sub UpdateMasterLists()

Dim arrGrade As Variant
Dim i As Integer
Dim tblMaster As ListObject
Dim tblGeneric As ListObject
Dim tblTBox As ListObject
Dim strGrade As String

Set tblGeneric = wsGenericActivities.ListObjects(1)
Set tblTBox = wsTBoxActivities.ListObjects(1)
arrGrade = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "DC3")
For i = 0 To UBound(arrGrade)
    strGrade = arrGrade(i)
    Set tblMaster = wsMasterList.ListObjects("tblMasterList" & strGrade)
    UpdateTblMaster tblMaster, tblGeneric, tblTBox, strGrade
Next

End Sub

Sub UpdateObjectiveAndDescription()

Dim tbl As ListObject
Dim strHeader As String
Dim strProjectN As String
Dim strActivityN As String
Dim r As Variant
Dim strActivity As String
Dim intRowIndex As Integer
Dim intColIndex As Integer
Dim r2 As Variant
Dim strSection As String
Dim strObjective As String
Dim strDescription As String

If Not ActiveCell.ListObject Is Nothing Then
    If LCase(ActiveCell.value) = "tbox" Then
        Set tbl = ActiveCell.ListObject
        strHeader = GetHeaderTitleFromMaster(ActiveCell)
        If strHeader = "Objective" Or strHeader = "Description" Then
            intRowIndex = ActiveCell.row - tbl.HeaderRowRange.row
            intColIndex = tbl.ListColumns("Actividad").Range.column - tbl.ListColumns(1).Range.column + 1
            strActivity = tbl.DataBodyRange(intRowIndex, intColIndex).value
            r = GetProjectAndActivityNumber(strActivity)
            strProjectN = r(0)
            strActivityN = r(1)
            strSection = GradeNameToNumeric(GetGradeSectionFromMaster(ActiveCell))
            r2 = GetTBoxProjectInfo(strProjectN, strActivityN, strSection)
            strObjective = r2(0)
            strDescription = r2(1)
            If strHeader = "Objective" Then
                ActiveCell.value = strObjective
            Else
                ActiveCell.value = strDescription
            End If
        End If
    End If
End If

End Sub

Sub UpdateTblMaster(tblMaster As ListObject, tblGeneric As ListObject, tblTBox As ListObject, strSection As String)
    Dim i As Long
    Dim strActivity As String
    Dim found As Boolean
    Dim rNumbers As Variant
    Dim strLessonNumber As String
    Dim strBimesterNumber As String
    Dim strProjectNumber As String
    
    ' Clear columns 4 and 5 from tblMaster
    tblMaster.ListColumns(3).DataBodyRange.ClearContents
    tblMaster.ListColumns(4).DataBodyRange.ClearContents
    tblMaster.ListColumns(5).DataBodyRange.ClearContents

    ' Loop through each row of column 2 in tblMaster
    For i = 1 To tblMaster.ListRows.count
        strActivity = tblMaster.DataBodyRange.Cells(i, 2).value ' Column 2
        rNumbers = GetNumbers(strActivity)
        strLessonNumber = rNumbers(1)
        strBimesterNumber = rNumbers(2)
        strProjectNumber = rNumbers(3)
        found = False ' Initialize found flag
        
        ' Try to find exact match in tblGeneric column 1
        Dim rngFound As Range
        Set rngFound = tblGeneric.DataBodyRange.Columns(1).Find(What:=strActivity, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not rngFound Is Nothing Then
            ' Found exact match
            tblMaster.DataBodyRange.Cells(i, 3).value = rngFound.Offset(0, 3).value
            tblMaster.DataBodyRange.Cells(i, 4).value = ReplaceNumberPlaceholders(rngFound.Offset(0, 1).value, _
                strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 2
            tblMaster.DataBodyRange.Cells(i, 5).value = ReplaceNumberPlaceholders(rngFound.Offset(0, 2).value, _
                strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 3
            found = True
        Else
            ' Replace numbers in strActivity with "<number>"
            Dim strActivityModified As String
            strActivityModified = ReplaceNumbersWithPlaceholder(strActivity)
            
            ' Try to find modified strActivity in tblGeneric
            Set rngFound = tblGeneric.DataBodyRange.Columns(1).Find(What:=strActivityModified, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not rngFound Is Nothing Then
                ' Found match
                tblMaster.DataBodyRange.Cells(i, 3).value = rngFound.Offset(0, 3).value
                tblMaster.DataBodyRange.Cells(i, 4).value = ReplaceNumberPlaceholders(rngFound.Offset(0, 1).value, _
                    strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 2
                tblMaster.DataBodyRange.Cells(i, 5).value = ReplaceNumberPlaceholders(rngFound.Offset(0, 2).value, _
                    strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 3
                found = True
            Else
                ' Check if strActivity matches "Project X, Activity Y"
                Dim projNum As String
                Dim actNum As String
                If ParseProjectActivity(strActivity, projNum, actNum) Then
                    ' Loop through tblTBox to find matching row
                    Dim j As Long
                    For j = 1 To tblTBox.ListRows.count
                        If tblTBox.DataBodyRange.Cells(j, 1).value = strSection Then ' Column 1
                            If tblTBox.DataBodyRange.Cells(j, 2).value = projNum Then ' Column 2
                                If tblTBox.DataBodyRange.Cells(j, 5).value = actNum Then ' Column 5
                                    ' Found match
                                    tblMaster.DataBodyRange.Cells(i, 3).value = tblTBox.DataBodyRange.Cells(j, 9).value
                                    tblMaster.DataBodyRange.Cells(i, 4).value = ReplaceNumberPlaceholders(tblTBox.DataBodyRange.Cells(j, 7).value, _
                                        strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 7
                                    tblMaster.DataBodyRange.Cells(i, 5).value = ReplaceNumberPlaceholders(tblTBox.DataBodyRange.Cells(j, 8).value, _
                                        strSection, strLessonNumber, strBimesterNumber, strProjectNumber) ' Column 8
                                    found = True
                                    Exit For
                                End If
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    Next i
End Sub

Function GetNumbers(strInput As String) As Variant
    Dim arrResult(1 To 3) As Variant
    Dim arrWords(1 To 3) As String
    Dim i As Integer
    Dim num As String

    ' Initialize array elements to empty string
    For i = 1 To 3
        arrResult(i) = ""
    Next i

    ' Define the words to search for
    arrWords(1) = "lesson"
    arrWords(2) = "bimester"
    arrWords(3) = "project"

    ' For each word, search in strInput
    For i = 1 To 3
        num = GetNumberAfterWord(LCase(strInput), arrWords(i))
        arrResult(i) = num
    Next i

    GetNumbers = arrResult
End Function

Function GetNumberAfterWord(strInput As String, word As String) As String
    Dim position As Long
    Dim afterWord As String
    Dim lenWord As Long
    Dim i As Long
    Dim lenAfter As Long
    Dim ch As String
    Dim num As String

    position = InStr(1, strInput, word, vbTextCompare)
    If position = 0 Then
        GetNumberAfterWord = ""
        Exit Function
    End If

    lenWord = Len(word)
    afterWord = Mid(strInput, position + lenWord)
    lenAfter = Len(afterWord)
    i = 1

    ' Skip any spaces or tabs after the word
    Do While i <= lenAfter
        ch = Mid(afterWord, i, 1)
        If ch = " " Or ch = vbTab Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    ' Collect digits
    num = ""
    Do While i <= lenAfter
        ch = Mid(afterWord, i, 1)
        If ch Like "[0-9]" Then
            num = num & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop

    GetNumberAfterWord = num
End Function
' Helper function to replace numbers with "<number>"
Function ReplaceNumbersWithPlaceholder(strInput As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "\d+"
        .Global = True
    End With
    ReplaceNumbersWithPlaceholder = regex.Replace(strInput, "<number>")
End Function

' Helper function to parse "Project X, Activity Y"
Function ParseProjectActivity(strInput As String, ByRef projNum As String, ByRef actNum As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "Project\s+(\d+),\s*Activity\s+(\d+)"
        .IgnoreCase = True
    End With
    Dim matches As Object
    Set matches = regex.Execute(strInput)
    If matches.count > 0 Then
        projNum = matches(0).SubMatches(0)
        actNum = matches(0).SubMatches(1)
        ParseProjectActivity = True
    Else
        ParseProjectActivity = False
    End If
End Function

Function GetGradeSectionFromMaster(ByVal rng As Excel.Range)

If Not rng.ListObject Is Nothing Then
    GetGradeSectionFromMaster = rng.ListObject.HeaderRowRange(1, 1).Offset(-1, 0).value
Else
    GetGradeSectionFromMaster = ""
End If

End Function

Function GetHeaderTitleFromMaster(ByVal rngCell As Excel.Range) As String

Dim intColIndex As Integer
Dim tbl As ListObject
Dim strHeader As String

strHeader = ""
If Not rngCell.ListObject Is Nothing Then
    Set tbl = rngCell.ListObject
    intColIndex = rngCell.column - tbl.HeaderRowRange(1, 1).column + 1
    strHeader = tbl.HeaderRowRange(1, intColIndex)
End If

GetHeaderTitleFromMaster = strHeader

End Function

Sub test()

Dim wordApp As word.Application

Set wordApp = New word.Application

wordApp.Visible = True

End Sub

' ===========================
' Health Check Integration
' ===========================

Sub RunHealthCheck()
    ' Main entry point for health check from main module
    RunBasicHealthCheck
End Sub

Sub TestHealthCheckMain()
    ' Test health check functionality
    TestHealthCheck
End Sub

Sub ShowGradebookInfoMain()
    ' Show gradebook structure information
    ShowGradebookInfo
End Sub

' ===========================
' Health Check for External Files
' ===========================

Sub RunHealthCheckOnFileMain()
    ' Run health check on a specific file
    Dim filePath As String
    filePath = InputBox("Enter the full path to the gradebook file:", "Health Check File")
    
    If filePath <> "" Then
        RunHealthCheckOnFile filePath
    End If
End Sub

Sub RunHealthCheckOnFolderMain()
    ' Run health check on all gradebooks in a folder
    Dim folderPath As String
    Dim bimester As String
    
    folderPath = InputBox("Enter the folder path containing gradebook files:", "Health Check Folder")
    If folderPath = "" Then Exit Sub
    
    bimester = InputBox("Enter bimester to filter (optional, leave blank for all):", "Health Check Filter")
    
    RunHealthCheckOnFolder folderPath, bimester
End Sub

Sub RunHealthCheckOnBimesterMain()
    ' Run health check on all week subfolders in a bimester folder
    Dim bimesterFolderPath As String
    
    bimesterFolderPath = InputBox("Enter the bimester folder path containing week subfolders (W01, W02, etc.):", "Health Check Bimester")
    If bimesterFolderPath = "" Then Exit Sub
    
    RunHealthCheckOnBimester bimesterFolderPath
End Sub

Sub RunHealthCheckOnCurrentWorkbookMain()
    ' Run health check on the currently active workbook
    RunHealthCheckOnWorkbook ActiveWorkbook
End Sub

Function GetUnitPlanContentsText(ByVal intGradeLevel As Integer, ByVal strProjectName As String) As String

' Composes a string from the Activity Number, Name, and Description for the corresponding intGradeLevel and strProjectName
' in listobject tblTBoxActivities in worksheet wsTBoxActivities.

Dim tbl As ListObject
Dim rngFiltered As Range
Dim intRow As Integer
Dim intActivityNumber As Integer
Dim strActivityName As String
Dim strActivityDescription As String
Dim strContent As String

Set tbl = wsTBoxActivities.ListObjects("tblTBoxActivities")

' Clear filter and show all data
If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData

' Filter by intGradeLevel
tbl.Range.AutoFilter Field:=1, Criteria1:="=" & intGradeLevel

' Filter by strProjectName
tbl.Range.AutoFilter Field:=3, Criteria1:="=" & strProjectName

' Get visible rows
Set rngFiltered = GetTableVisibleRows(tbl)

' Get composed string
strContent = ""
If Not rngFiltered Is Nothing Then
    For intRow = 1 To rngFiltered.Rows.count
        intActivityNumber = rngFiltered.Cells(intRow, 5)
        strActivityName = rngFiltered.Cells(intRow, 6)
        strActivityDescription = rngFiltered.Cells(intRow, 8)
        If strActivityDescription <> "" Then
            strContent = strContent & "Activity " & intActivityNumber & ": " & _
                strActivityName & vbCrLf & vbCrLf & strActivityDescription
            If intRow < rngFiltered.Rows.count Then
                strContent = strContent & vbCrLf & vbCrLf
            End If
        Else
            strContent = ""
        End If
    Next
End If

GetUnitPlanContentsText = strContent

End Function

Function GetUnitPlanObjectivesText(ByVal intGradeLevel As Integer, ByVal strProjectName As String) As String

' Composes a string from the Activity Objectives for the corresponding intGradeLevel and strProjectName
' in listobject tblTBoxActivities in worksheet wsTBoxActivities.

Dim tbl As ListObject
Dim rngFiltered As Range
Dim intRow As Integer
Dim intActivityNumber As Integer
Dim strActivityName As String
Dim strActivityObjective As String
Dim strContent As String

Set tbl = wsTBoxActivities.ListObjects("tblTBoxActivities")

' Clear filter and show all data
If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData

' Filter by intGradeLevel
tbl.Range.AutoFilter Field:=1, Criteria1:="=" & intGradeLevel

' Filter by strProjectName
tbl.Range.AutoFilter Field:=3, Criteria1:="=" & strProjectName

' Get visible rows
Set rngFiltered = GetTableVisibleRows(tbl)

' Get composed string
strContent = ""
If Not rngFiltered Is Nothing Then
    For intRow = 1 To rngFiltered.Rows.count
        intActivityNumber = rngFiltered.Cells(intRow, 5)
        strActivityName = rngFiltered.Cells(intRow, 6)
        strActivityObjective = rngFiltered.Cells(intRow, 7)
        If strActivityObjective <> "" Then
            strContent = strContent & "Activity " & intActivityNumber & ": " & _
                strActivityObjective
            If intRow < rngFiltered.Rows.count Then
                strContent = strContent & vbCrLf & vbCrLf
            End If
        Else
            strContent = ""
        End If
    Next
End If

GetUnitPlanObjectivesText = strContent

End Function


Function GetTableVisibleRows(lo As ListObject) As Range
    Dim r As Range

    On Error Resume Next ' In case table has no data rows
    Set r = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    Set GetTableVisibleRows = r
End Function

Sub ImportTBoxProjectData()
    Dim folderPath As String
    Dim fso As Object, file As Object
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim jsonText As String
    Dim json As Object
    Dim activitiesText As String
    Dim standardsText As String
    Dim indicatorsText As String
    Dim fileName As String, projectName As String
    Dim gradeText As String, gradeValue As Variant
    Dim projectExists As Boolean
    Dim lineBreak As String: lineBreak = Chr(10)
    Dim intProjectNumber As Integer
    Dim strContentsText As String
    Dim strObjectivesText As String
    
    ' Set worksheet and table
    Set ws = wsTBoxProjectsInfo
    Set tbl = ws.ListObjects("tblTBoxProjectsInfo")
    
    ' Get folder path
    folderPath = GetOneDriveRoot() & "\2526\Computers\TBox Projects\Processed\"
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' FileSystemObject to loop through files
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        MsgBox "Folder not found: " & folderPath, vbExclamation
        Exit Sub
    End If
    
        'Declare the ProgressBar Object
    Dim MyProgressbar As ProgressBar
    'Initialize a New Instance of the Progressbars
    Set MyProgressbar = New ProgressBar
    
    'Set all the Properties that need to be set before the
    'ProgresBar is Shown
    With MyProgressbar
        'Set the Title
        .Title = "Updating TBox Projects Info"
        'Set this to true if you want to update
        'Excel's Status Bar Also
        .ExcelStatusBar = True
        'Set the colour of the bar in the Beginning
        .StartColour = rgbMediumSeaGreen
        'Set the colour of the bar at the end
        .EndColour = rgbGreen
        .TotalActions = fso.GetFolder(folderPath).Files.count 'Required
    End With
    
    'Show the Bar
    MyProgressbar.ShowBar 'Critical Line

    
    For Each file In fso.GetFolder(folderPath).Files
        'Update the ProgressBar NextAction Method
        MyProgressbar.NextAction "Processing '" & file.Name & "'", True
        
        If LCase(fso.GetExtensionName(file.Name)) = "txt" Then
            fileName = fso.GetBaseName(file.Name)
            
            ' Parse project number
            intProjectNumber = GetProjectNumber(file.Name)
            
            ' Parse grade and project from filename
            gradeText = ExtractGrade(fileName)
            gradeValue = GradeTextToNumber(gradeText)
            If IsEmpty(gradeValue) Then GoTo NextFile
            
            ' Check if already exists
            projectExists = False
            Dim r As ListRow
            For Each r In tbl.ListRows
                If r.Range(1, 1).value = fileName And r.Range(1, 3).value = gradeValue Then
                    projectExists = True
                    Exit For
                End If
            Next r
            If projectExists Then GoTo NextFile
            
            ' Read JSON
            Set json = ParseJsonFileClean(file.path)
            
            ' Project name
            If json.Exists("project_title") Then
                projectName = json("project_title")
            End If
            
            ' Activities
            Dim act As Object
            activitiesText = ""
            If json.Exists("activities") Then
                For Each act In json("activities")
                    activitiesText = activitiesText & act("title") & ": " & act("purpose") & lineBreak
                Next act
                activitiesText = Left(activitiesText, Len(activitiesText) - Len(lineBreak)) ' remove last line break
            End If
            
            ' Standards
            standardsText = ""
            
            If json.Exists("iste_standards") Then
                Select Case TypeName(json("iste_standards"))
                
                    Case "Collection"
                        Dim itm As Variant
                        For Each itm In json("iste_standards")
                            standardsText = standardsText & itm & lineBreak
                        Next itm
                        'remove trailing break
                        If Len(standardsText) > 0 Then _
                            standardsText = Left$(standardsText, Len(standardsText) - Len(lineBreak))
                            
                    Case Else 'array or single value
                        If IsArray(json("iste_standards")) Then
                            standardsText = Join(json("iste_standards"), lineBreak)
                        Else
                            standardsText = CStr(json("iste_standards"))
                        End If
                                
                End Select
            End If
            
            ' Indicators
            indicatorsText = GetFormattedIndicators(json)
            
            ' Contents
            strContentsText = GetUnitPlanContentsText(gradeValue, projectName)
            
            ' Objectives
            strObjectivesText = GetUnitPlanObjectivesText(gradeValue, projectName)
            
            ' Add new row
            Set newRow = tbl.ListRows.Add
            With newRow.Range
                .Cells(1, 1).value = intProjectNumber ' Project #
                .Cells(1, 2).value = projectName ' Project Name
                .Cells(1, 3).value = gradeValue ' Grade
                .Cells(1, 6).value = activitiesText ' Activities
                .Cells(1, 7).value = standardsText ' Standards
                .Cells(1, 8).value = strContentsText ' Contents
                .Cells(1, 9).value = strObjectivesText ' Objectives
                .Cells(1, 10).value = indicatorsText ' Indicators
            End With
            
        End If
NextFile:
    Next file
    
    ' Date Start
    With tbl.ListColumns(4).DataBodyRange
        .formula = "=XLOOKUP(TRUE, ISNUMBER(SEARCH([@[Project Name]],tblProjects[Project])),tblProjects[Date Start])"
        .NumberFormat = "dd-mmm-yyyy"
    End With
    
    ' Date End
    With tbl.ListColumns(5).DataBodyRange
        .formula = "=XLOOKUP(TRUE, ISNUMBER(SEARCH([@[Project Name]],tblProjects[Project])),tblProjects[Date End])"
        .NumberFormat = "dd-mmm-yyyy"
    End With
    
    MyProgressbar.Complete
    
    MsgBox "Import completed.", vbInformation
End Sub

Function GetFormattedIndicators(json As Object) As String
    Dim indicatorsDict As Object
    Dim key As Variant
    Dim indicator As Variant
    Dim result As String
    
    Set indicatorsDict = json("indicators")
    
    For Each key In indicatorsDict.Keys
        result = result & key & vbCrLf
        For Each indicator In indicatorsDict(key)
            result = result & "  " & indicator & vbCrLf
        Next indicator
        result = result & vbCrLf
    Next key
    
    GetFormattedIndicators = result
End Function

Function GetProjectNumber(ByVal strFileName As String) As Integer
    Dim regex As Object
    Dim match As Object
    Dim matches As Object

    ' Create regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = "Project\s+(\d+)"
        .Global = False
        .IgnoreCase = True
    End With

    ' Execute the regex
    Set matches = regex.Execute(strFileName)
    If matches.count > 0 Then
        Set match = matches(0)
        GetProjectNumber = CInt(match.SubMatches(0))
    Else
        GetProjectNumber = 0 ' Return 0 if no match found
    End If
End Function

Function GradeTextToNumber(gradeText As String) As Variant
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")
    map.Add "First", 1
    map.Add "Second", 2
    map.Add "Third", 3
    map.Add "Fourth", 4
    map.Add "Fifth", 5
    map.Add "Sixth", 6
    map.Add "Seventh", 7
    map.Add "Eighth", 8
    map.Add "Ninth", 9
    map.Add "Tenth", 10
    map.Add "Eleventh", 11
    map.Add "Twelfth", 12
    
    If map.Exists(gradeText) Then
        GradeTextToNumber = map(gradeText)
    Else
        GradeTextToNumber = Empty
    End If
End Function
