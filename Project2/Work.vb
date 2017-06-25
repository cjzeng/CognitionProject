'Module Work

Sub doIt()
    Dim converter As New IngestTool
    converter.processAll
End Sub

Sub processWorkbook(filePath As String)
    Dim activeWk As String
    
    'save the current workbook name
    activeWk = ActiveWorkbook.name
    
    'open the workbook
    Workbooks.Open filename:=fi
    'process the worksheet
    'close the workbook
End Sub

Sub testfindDimention()
    Dim converter As New IngestTool
    'converter.findDimention
    converter.processAll
End Sub

Sub testGetSection()
    Dim converter As New IngestTool
    'converter.findDimention
    Debug.Print converter.getSection("Lecture 1.1 9-6-2016 4-41 PM Report.xlsx")
    Debug.Print converter.getSection("Lecture 1.2 9-6-2016 4-41 PM Report.xlsx")
End Sub

Sub testconvertToLetter()
    Dim converter As New IngestTool
    'converter.findDimention
    Debug.Print converter.convertToLetter("A")
    Debug.Print converter.convertToLetter("B")
    Debug.Print converter.convertToLetter("C")
    Debug.Print converter.convertToLetter("D")
    Debug.Print converter.convertToLetter("E")
    Debug.Print converter.convertToLetter("-")
    Debug.Print converter.convertToLetter("Aldf ddfd")
End Sub
