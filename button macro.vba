Sub Button1_Click()

    Dim lastRow As Long
    Dim i As Long
    
    ' Find the last row with data in Column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in Column B where there is a corresponding value in Column F
    For i = 2 To lastRow
        If Not IsEmpty(Cells(i, 6)) Then ' Check if there's a value in Column B
            Cells(i, 2).Value = Cells(i, 2).Value + Cells(i, 6).Value
            Cells(i, 6).ClearContents ' Clear the value in Column B after adding
        End If
    Next i

End Sub
