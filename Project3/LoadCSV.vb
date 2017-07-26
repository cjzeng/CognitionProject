
Sub list_files()
   Dim MyObj As Object, MySource As Object, file As Variant, count As Integer, limit As Integer
   file = Dir("C:\Users\Daddy\Documents\DrArnoldGlass\CognitionProject2\")
   count = 1
   limit = 500
   Do While (file <> "")
      'Debug.Print "found file=" & file
      If InStr(file, ".csv") > 0 Then
           Debug.Print "found csv file=" & file
           import_file file
           count = count + 1
      End If
      If (count > limit) Then
        Exit Do
      End If
        
      file = Dir
      'Debug.Print "next file=", file
    Loop
End Sub

Public Sub import_files()
    Dim baseDir As String
    baseDir = "/App/Data/Downloads/cognitionproject/"
   import_file baseDir & "testdata 1.0.csv"
   import_file baseDir & "testdata 1.0_p.csv"
   import_file baseDir & "testdata 10.1.csv"
   import_file baseDir & "testdata 10.2.csv"


End Sub

Sub import_file(filename As Variant)
    import_file_to_worksheet filename
    remove_non_data_rows
    rename_column_headers_delete_data_column
End Sub


Sub import_file_to_worksheet(filename As Variant)
'    Dim lIdx As Long
'    Dim rIdx As Long

'    lIdx = InStrRev(filename, "/")
'    rIdx = InStrRev(filename, ".")
'    Dim sheetName
'    sheetName = Mid(filename, lIdx + 1, rIdx - lIdx - 1)
'    Debug.Print "adding data sheet:", sheetName, "-", filename
'    Sheets.Add.Name = sheetName
'    Sheets(sheetName).Select
    import_csv_file filename
End Sub


Sub import_csv_file(filename As Variant)
  'Position As Range
  'Position = ActiveSheet.Cells(1, 1)
      Dim lIdx As Long
    Dim rIdx As Long

    lIdx = InStrRev(filename, "/")
    rIdx = InStrRev(filename, ".")
    Dim sheetName
    sheetName = Mid(filename, lIdx + 1, rIdx - lIdx - 1)
    Debug.Print "adding data sheet:", sheetName, "-", filename
    'Sheets.Add.Name = sheetName
    ActiveWorkbook.Worksheets.Add.Name = sheetName
    With ActiveSheet.QueryTables.Add(Connection:= _
        "TEXT;" & filename _
        , Destination:=ActiveSheet.Cells(1, 1))
      .Name = Replace(filename, ".csv", "")
      .FieldNames = True
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 1252
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
        1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

Sub remove_non_data_rows()
    Dim rowSize As Long
    Dim r As Long
    Dim emptyColumn As Long
    emptyColumn = 10
'    rowSize = ActiveSheet.UsedRange.Rows.Count + 1
'    rowSize = Cells(ActiveSheet.Rows.count, 1).End(xlUp).row
    Debug.Print "rowSize=", rowSize
    For r = rowSize To 1 Step -1
        'Rows(r).Select
        Set cellRange = ActiveSheet.Range(ActiveSheet.Cells(r, emptyColumn), ActiveSheet.Cells(r, emptyColumn))
        If IsEmpty(cellRange) Then
            'cellRange.Select
            ActiveSheet.Rows(r).Delete
'            Debug.Print "Deleting row =", r
        End If
    Next
End Sub

Sub rename_column_headers_delete_data_column()

    'Find the row with name value
    Dim rowSize As Long
    Dim nameColumn As Long
    Dim r As Long
    Dim idName As String

    nameColumn = 3
    idName = "name"

    rowSize = Cells(ActiveSheet.Rows.count, 1).End(xlUp).row
    For r = rowSize To 1 Step -1
        'ActiveSheet.Rows(r).Select
'        Debug.Print r, nameColumn, Cells(r, nameColumn).Value
        If Cells(r, nameColumn).Value = idName Then
'            Cells(r, nameColumn).Select
'            Debug.Print "row=", r
            rename_headers r
            remove_data_columns r
            Exit For
        End If
    Next

End Sub

Sub rename_headers(row As Long)

    Dim colSize As Long
    Dim cellValue
    
    colSize = ActiveSheet.UsedRange.Columns.count
'    Debug.Print "colSize=", colSize
    For r = 1 To colSize
'        Cells(row, r).Select
        cellValue = Cells(row, r).Value
'        Debug.Print row, r, cellValue
        If (InStr(cellValue, ":")) Then
'            Debug.Print "found", cellValue
            Dim lIdx As Integer
            Dim rInx As Integer
            lIdx = InStr(cellValue, "Q")
            If (lIdx = 0) Then
                lIdx = InStr(cellValue, "D")
            End If
            rIdx = InStr(cellValue, ".")
            If (rIdx - lIdx > 0) Then
                Dim newHeader As String
                newHeader = Mid(cellValue, lIdx, rIdx - lIdx)
                Debug.Print "newHeader=", newHeader, " from older header:", cellValue, lIdx, rIdx
                Cells(row, r).Value = newHeader
            End If
        End If
    Next

End Sub

Sub remove_data_columns(row As Long)

    Dim colSize As Long
    Dim cellValue
    
    colSize = ActiveSheet.UsedRange.Columns.count
'    Debug.Print "colSize=", colSize
    For r = colSize To 1 Step -1
'        Cells(row, r).Select
        cellValue = Cells(row, r).Value
        Debug.Print row, r, cellValue
        If ((InStr(cellValue, "D") = 0) And (InStr(cellValue, "Q") = 0) And (InStr(cellValue, "name") = 0)) Then
'            Debug.Print "deleting column ", r, " with ", cellValue
            ActiveSheet.Columns(r).Delete
        End If
    Next

End Sub


