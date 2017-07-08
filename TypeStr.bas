Sub TypeStr()

Dim I As Long
Dim LastRow As Long

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To LastRow
Cells(I, 1).Value = CStr(Cells(I, 1).Value)
Next I
   
End Sub