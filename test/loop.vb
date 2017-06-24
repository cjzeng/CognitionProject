Sub walk_cell()
    Dim rowSize As Long
    Dim r As Long
    Dim colSize As Long
    Dim c As Long
   
    rowSize = 20
    colSize = 20
    For r = 14 To 54
        ActiveSheet.Rows(r).Select
'        Debug.Print r, nameColumn, Cells(r, nameColumn).Value
        For c = 1 To 7
            Cells(r, c).Select
            Debug.Print "row=", r, ", col=", c, ", value=", Cells(r, c).Value
            Cells(r + 10, c + 10).Value = Cells(r, c).Value
            Cells(r + 10, c + 10).Select
        Next c
    Next
End Sub
