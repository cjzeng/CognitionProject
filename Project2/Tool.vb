'Module Tool

Sub load_files_in_dir(folder As String, suffix As String)
   Dim MyObj As Object, MySource As Object, file As Variant
   file = dir(folder & suffix)
   While (file <> "")
      Debug.Print "loading file=", file
      Workbooks.Open filename:=file
    Wend
End Sub

Sub list_files_in_dir(folder As String, suffix As String)
   Dim MyObj As Object, MySource As Object, file As Variant
   file = dir(folder & suffix)
   While (file <> "")
      Debug.Print "found file=", file
      If InStr(file, ".csv") > 0 Then
           Debug.Print "found csv file=", file
      End If
      file = dir
'      Debug.Print "next file=", file
    Wend
End Sub

Sub list_files()
   list_files_in_dir "D:\Clicker Responses\", "*.xslx"
End Sub

Sub load_files()
   load_files_in_dir "D:\Clicker Responses\", "*.xslx"
End Sub

Sub ListWorkbooks()
    Dim Rng As Range
    Dim WorkRng As Range
    On Error Resume Next
    xTitleId = "KutoolsforExcel"
    Set WorkRng = Application.Selection
    Set WorkRng = Application.InputBox("Out put to (single cell)", xTitleId, WorkRng.Address, Type:=8)
    Set WorkRng = WorkRng.Range("A1")
    xNum1 = Application.Workbooks.Count
    For i = 1 To xNum1
        xNum2 = Application.Workbooks(i).Sheets.Count
        WorkRng.Offset(i - 1, 0).value = Application.Workbooks(i).name
        For j = 1 To xNum2
            WorkRng.Offset(i - 1, j).value = Application.Workbooks(i).Sheets(j).name
        Next
    Next
End Sub

Sub list_work_books()
    Dim wbCount As Integer
    Dim sheet As Worksheet
    Dim book As Workbook
    Dim bookname As String
    
    wbCount = Application.Workbooks.Count
    
    For i = 1 To wbCount
        bookname = Application.Workbooks(i).name
        Application.Workbooks(i).Activate
        For j = 1 To Application.Workbooks(i).Sheets.Count
            Debug.Print bookname, ", ", Application.Workbooks(i).Sheets(j).name
        Next
        If (InStr(bookname, ".xlsx") > 0) Then
            Debug.Print "found ", bookname
        End If
    Next


End Sub

Sub CopyWorkbook()

    Dim currentSheet As Worksheet
    Dim sheetIndex As Integer
    sheetIndex = 1

    For Each currentSheet In Worksheets

        Windows("SOURCE WORKBOOK").Activate
        currentSheet.Select
        currentSheet.Copy Before:=Workbooks("TARGET WORKBOOK").Sheets(sheetIndex)

        sheetIndex = sheetIndex + 1

    Next currentSheet

End Sub
