Attribute VB_Name = "Tools"
Sub OpenRichTextEditor()
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim userInput As String
    
    ' Create a new instance of Word
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set wordApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    ' Make Word visible and create a new document
    wordApp.Visible = True
    Set wordDoc = wordApp.Documents.Add
    
    ' Prompt the user to enter rich text in Word
    MsgBox "Please enter your rich text in the Word document, then close it to continue.", vbInformation, "Rich Text Editor"
    
    ' Wait until the document is closed
    Do While wordDoc.Saved = False Or wordApp.Documents.Count > 0
        DoEvents
    Loop
    
    ' Copy the content from Word
    wordDoc.Content.Copy
    
    ' Paste it into the active cell in Excel
    ActiveCell.PasteSpecial xlPasteValues
    
    ' Close Word without saving
    wordApp.Quit False
    
    MsgBox "Rich text has been inserted into the Excel cell.", vbInformation, "Complete"
End Sub


Sub AutoFitColumnsInListObjects()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim maxWidth As Double
    Dim colWidth As Double
    
    maxWidth = 30 ' Maximum allowed width for columns

    ' Set the active sheet as the worksheet to operate on
    Set ws = ActiveSheet
    
    ' Loop through all list objects (tables) in the active sheet
    For Each lo In ws.ListObjects
        ' Loop through all columns in the list object
        For Each lc In lo.ListColumns
            ' Autofit the column
            lc.Range.Columns.AutoFit
            
            ' Check the new width of the column
            colWidth = lc.Range.Columns(1).ColumnWidth
            
            ' If the width exceeds the maximum allowed width
            If colWidth > maxWidth Then
                ' Set the column width to the maximum allowed width
                lc.Range.Columns(1).ColumnWidth = maxWidth
                
                ' Enable wrap text for the column
                lc.Range.Cells.WrapText = True
            End If
        Next lc
    Next lo
End Sub

