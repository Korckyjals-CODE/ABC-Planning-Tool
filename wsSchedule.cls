VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' Show/hide document icon based on selected cell content
    Dim tblSchedule As ListObject
    Dim rngSelected As Range
    Dim strCellValue As String
    
    ' First, hide all existing icons since selection has changed
    HideAllDocumentIcons
    
    ' Get the tblSchedule listobject
    Set tblSchedule = Me.ListObjects("tblSchedule")
    
    ' Check if the selection is within the tblSchedule data range
    If Not Intersect(Target, tblSchedule.DataBodyRange) Is Nothing Then
        Set rngSelected = Target.Cells(1, 1) ' Handle multi-cell selection
        strCellValue = CStr(rngSelected.Value)
        
        ' Show document icon only if cell content matches grade level pattern
        If IsGradeLevel(strCellValue) Then
            ShowDocumentIcon rngSelected
        End If
    End If
End Sub

Private Function IsGradeLevel(ByVal strValue As String) As Boolean
    ' Check if the string matches grade level patterns: 1A-12B, DC3A, DC3B
    Dim grade As Integer
    Dim section As String
    
    ' Handle special cases first
    If strValue = "DC3A" Or strValue = "DC3B" Then
        IsGradeLevel = True
        Exit Function
    End If
    
    ' Check if it's a standard grade level (1A-12B)
    If Len(strValue) = 2 Then
        If IsNumeric(Left(strValue, 1)) Then
            grade = CInt(Left(strValue, 1))
            section = Right(strValue, 1)
            If grade >= 1 And grade <= 12 And (section = "A" Or section = "B") Then
                IsGradeLevel = True
                Exit Function
            End If
        End If
    ElseIf Len(strValue) = 3 Then
        ' Handle double-digit grades (10A-12B)
        If IsNumeric(Left(strValue, 2)) Then
            grade = CInt(Left(strValue, 2))
            section = Right(strValue, 1)
            If grade >= 10 And grade <= 12 And (section = "A" Or section = "B") Then
                IsGradeLevel = True
                Exit Function
            End If
        End If
    End If
    
    IsGradeLevel = False
End Function

Private Sub ShowDocumentIcon(ByVal rngCell As Range)
    ' Show a small document icon to the right of the selected cell
    Dim rngIcon As Range
    Dim shpIcon As Shape
    
    ' Define the position for the icon (one cell to the right)
    Set rngIcon = rngCell.Offset(0, 1)
    
    ' Remove any existing icon in this position
    HideDocumentIcon rngCell
    
    ' Add a small document icon
    Set shpIcon = Me.Shapes.AddShape(msoShapeRectangle, _
        rngIcon.Left + 2, rngIcon.Top + 2, 16, 16)
    
    ' Style the icon to look like a document
    With shpIcon
        .Name = "DocIcon_" & rngCell.Address(False, False)
        .Fill.ForeColor.RGB = RGB(255, 255, 255) ' White background
        .Line.ForeColor.RGB = RGB(0, 0, 0) ' Black border
        .Line.Weight = 1
        .OnAction = "DocumentIcon_Click" ' Assign click event
    End With
    
    ' Add text to make it look like a document
    With shpIcon.TextFrame
        .Characters.Text = "📄"
        .Characters.Font.Size = 10
        .Characters.Font.Name = "Segoe UI Emoji"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub HideDocumentIcon(ByVal rngCell As Range)
    ' Hide the document icon for the specified cell
    Dim shpIcon As Shape
    Dim iconName As String
    
    iconName = "DocIcon_" & rngCell.Address(False, False)
    
    On Error Resume Next
    Set shpIcon = Me.Shapes(iconName)
    If Not shpIcon Is Nothing Then
        shpIcon.Delete
    End If
    On Error GoTo 0
End Sub

Private Sub HideAllDocumentIcons()
    ' Hide all document icons on the worksheet
    Dim shp As Shape
    Dim i As Integer
    
    ' Loop backwards through shapes to avoid index issues when deleting
    For i = Me.Shapes.Count To 1 Step -1
        Set shp = Me.Shapes(i)
        If shp.Name Like "DocIcon_*" Then
            shp.Delete
        End If
    Next i
End Sub

Public Sub DocumentIcon_Click()
    ' Dummy message box for document icon click
    ' This will be replaced with actual functionality later
    MsgBox "Document icon clicked! This is a placeholder message.", vbInformation, "Document Icon"
End Sub