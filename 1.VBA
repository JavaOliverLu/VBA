Sub AA()
 Sheets(1).Name = "new"
 Dim s, b As Integer
 For a = 1 To 5
    For b = a To 5
    Sheets("new").Cells(a, b) = "*"
    Next
 Next
End Sub
