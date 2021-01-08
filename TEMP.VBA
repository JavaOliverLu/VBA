Sub AA()
 Sheets(1).Name = "new"
 Dim s, b As Integer
 For a = 1 To 3
    For b = 1 To 4
    Sheets("new").Cells(a + 5, b) = (Sheets("new").Cells(a, b) + 271)
    Sheets("new").Cells(a + 10, b) = (1.8 * (Sheets("new").Cells(a, b)) + 32)
    
    Next
 Next
 Sheets("new").Cells(5, 1) = "克氏溫標"
 Sheets("new").Cells(10, 1) = "華氏溫標"
End Sub
