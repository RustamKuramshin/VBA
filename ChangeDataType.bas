Attribute VB_Name = "ChangeDataType"

Sub Macros()

'Макрос Excel

Dim nameList As String

For Each lst In ThisWorkbook.Worksheets

nameList = lst.Name

Sheets(nameList).Activate

For i = 6 To 28
    For j = 19 To 30
        If InStr(Cells(i, j).Formula, "*") = 0 Then
            Cells(i, j).Value = CDbl(Cells(i, j).Value)
            Cells(i, j).NumberFormat = "0.00"
        End If
    Next
Next

For i = 6 To 28
    For j = 19 To 30
        If TypeName(Cells(i, j).Value) <> "Double" Then
            MsgBox "Ошибка в строке", i
        End If
    Next
Next

Next

End Sub
