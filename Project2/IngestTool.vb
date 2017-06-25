Option Explicit

Dim fileList(1 To 200) As String
Dim listSize As Integer
Dim folder As String
Dim suffix As String
Dim lookup As New NameLookUp
Dim dest As New dest
Dim src As New SourceMap

Dim bRow As Long
Dim eRow As Long
Dim colWidth As Long



Public Sub processAll()
    Dim i As Long
    loadFileList
    For i = 1 To listSize
        processWorkbook getFilePath(i)
    Next
    Debug.Print "Done ---------------"
    
'    processWorkbook getFilePath(1)
'    processWorkbook getFilePath(2)
'    processWorkbook getFilePath(3)
'    processWorkbook getFilePath(4)
End Sub

Function loadFileList() As Integer

   Dim MyObj As Object, MySource As Object, file As Variant
   file = dir(folder & suffix)
   While (file <> "")
      listSize = listSize + 1
      fileList(listSize) = file
      Debug.Print listSize, file
      file = dir
    Wend
    loadFileList = listSize
End Function

Function getFilePath(idx As Long) As String
   getFilePath = fileList(idx)
End Function

Function getSection(filePath As String) As String
    Dim i, j, k As Integer
    Dim section As String
    
    section = Split(filePath, " ")(1)
    getSection = Split(section, ".")(1)
End Function
Sub processWorkbook(filePath As String)
    Dim activeWk As String
    
    'save the current workbook name
    activeWk = ActiveWorkbook.name
    
    'open the workbook
    Workbooks.Open filename:=folder & filePath
    
    'process the worksheet
    Debug.Print ActiveWorkbook.name, " old ="; activeWk
'    processWorksheet
    copyToSource getSection(filePath)
    
    'close the workbook
    Workbooks(filePath).Close SaveChanges:=False
    
    Application.Workbooks(activeWk).Activate
    Application.Workbooks(activeWk).Sheets(1).Activate
    
    pasteToDestination

    
End Sub

Sub findDimention()
'    rowSize = ActiveSheet.UsedRange.Rows.Count + 1
    Dim rowSize As Long
    Dim startCol As Long
    Dim r As Long
    
    rowSize = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).row
    colWidth = ActiveSheet.UsedRange.Columns.Count
    Debug.Print "rowSize=", rowSize, ", colWidth=", colWidth
    
    startCol = 1
    For r = 1 To rowSize
        If (Cells(r, startCol).Text = "First Name") Then
            bRow = r
        End If
        If (Cells(r, startCol).Text = "Participant List Averages") Then
            eRow = r
        End If
    Next

    Debug.Print bRow & ", eRow=" & eRow & ", colWidth=" & colWidth
    
    
End Sub
Sub copyToSource(section As String)
    Debug.Print ActiveWorkbook.Worksheets(1).name
    Dim i As Long
    Dim dataLength As Long
    
    findDimention
    
    src.init
    'build header
    src.setSection section
    
    For i = 1 To colWidth
        If (Cells(bRow, i).Text <> "") Then
            src.setColumnName i, Cells(bRow, i).Text
        End If
    Next
    src.buildDataMap
    
    'set up destination
    dataLength = src.getDataMapSize
    dest.buildColumnMap src.getFieldMap, dataLength
    
'    For i = 1 To dest.getCMapSize()
'        Debug.Print "colmap ", i, "=", dest.getColumnMapItem(i)
'    Next
'
'    For i = 1 To dataLength
'        Debug.Print "dataMap ", i, "=", dest.getDataMapIdx(i)
'    Next

    'Read data rows
    Dim r As Long
    Dim c As Long
    Dim value As String
    
    r = 1
    For i = bRow + 1 To eRow - 1
        If (Cells(i, 1).Text <> "") Then
            For c = 1 To dataLength
                value = Cells(i, src.dataMapItem(c)).Text
                src.setDataItem r, c, convertToLetter(value)
                'Debug.Print "r=", r, ", c=", c, ", val=", Cells(i, src.dataMapItem(c)).Text, "(", i, ",", src.dataMapItem(c), ")"
            Next
            r = r + 1
        End If
    Next
End Sub

Function convertToLetter(value As String) As String
    Dim str As String
    
    str = value
    
    Select Case value
        Case Is = "A"
            str = "1"
        Case Is = "B"
            str = "2"
        Case Is = "C"
            str = "3"
        Case Is = "D"
            str = "4"
        Case Is = "E"
            str = "5"
        Case Is = "-"
            str = ""
        Case Else
            str = value
     End Select
    convertToLetter = str
End Function

Sub pasteToDestination()

    Dim r As Long
    Dim c As Long
    
    'Copy column header row if any
    For c = 1 To dest.getCMapSize
        Cells(1, c).Select
        Cells(1, c).value = dest.getColumnMapItem(c)
    Next
    'Copy answer row
    For c = 1 To src.getDataMapSize
        Cells(2, dest.getDataMapIdx(c)).Select
        Cells(2, dest.getDataMapIdx(c)).value = src.getDataItem(1, c)
    Next
    'Copy data row
    For r = 2 To src.getDataItemSize
        copyDataRow (r)
    Next
End Sub

Sub copyDataRow(srcRow As Long)
    Dim destRow As Long
    Dim nameKey As String
    Dim c As Long
    
    nameKey = src.getDataItem(srcRow, 1) & src.getDataItem(srcRow, 2)
    destRow = lookup.getNameIdx(nameKey) + 2
    Debug.Print "section=" & src.getSection & " , destRow=" & destRow & ", nameKey=" & nameKey
    
    For c = 1 To src.getDataMapSize
        Cells(destRow, dest.getDataMapIdx(c)).value = src.getDataItem(srcRow, c)
        'Cells(destRow, dest.getDataMapIdx(c)).Select
    Next
    Cells(destRow, 3).value = src.getSection
End Sub

Private Sub Class_Initialize()
    folder = "D:\Clicker Responses\"
    suffix = "*.xlsx"
    dest.init
End Sub
