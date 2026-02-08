Attribute VB_Name = "VeyonMessages"
Option Explicit

Private Function JsonEscape(ByVal value As String) As String
    ' Minimal JSON escaping for common characters
    value = Replace(value, "\", "\\")
    value = Replace(value, """", "\""")
    value = Replace(value, Chr(8), "\b")
    value = Replace(value, Chr(9), "\t")
    value = Replace(value, Chr(10), "\n")
    value = Replace(value, Chr(12), "\f")
    value = Replace(value, Chr(13), "\r")
    JsonEscape = value
End Function

Public Sub ExportMessagesToJson()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim pc As String, msg As String
    Dim json As String
    Dim first As Boolean
    Dim filePath As String
    Dim f As Integer

    ' Change "Messages" if your sheet has a different name
    Set ws = ThisWorkbook.Worksheets("VeyonMessages")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    json = "{"
    first = True

    For i = 2 To lastRow   ' Row 1 = headers
        pc = Trim(ws.Cells(i, "A").value)
        msg = CStr(ws.Cells(i, "B").value)

        If pc <> "" Then
            If Not first Then
                json = json & "," & vbCrLf
            Else
                json = json & vbCrLf
                first = False
            End If

            json = json & "  """ & pc & """: """ & JsonEscape(msg) & """"
        End If
    Next i

    json = json & vbCrLf & "}"

    ' Save Messages.json in the same folder as the workbook
    filePath = ThisWorkbook.path & "\Messages.json"

    f = FreeFile
    Open filePath For Output As #f
    Print #f, json
    Close #f

    MsgBox "Messages.json exported to:" & vbCrLf & filePath, vbInformation, "Export complete"
End Sub

