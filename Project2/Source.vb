Option Explicit

'Dim maxArraySize As Long

Const maxArraySize = 400
Const maxColumnSize = 30
'columnName - maps index to header
Dim cMax As Long  'columnName actaul size used
Dim columeName(1 To maxArraySize) As String 'columnName array

'dataMap - maps of indexs to usable data column
Dim dataMap(1 To maxArraySize) As Long
Dim dataMapSize As Long 'The length of dataMap array

'fieldMap - field name for each index in data map, uses dataMapSize for actual size
Dim fieldMap(1 To maxArraySize) As String

Dim dataRows(1 To maxArraySize, 1 To maxColumnSize) As String
Dim dRowSize As Long
Dim section As String

Sub init()
    cMax = 0
    dataMapSize = 0
    dRowSize = 0
    Dim i As Long
    Dim j As Long
    For i = 1 To maxArraySize
        columeName(i) = ""
        dataMap(i) = 0
        fieldMap(i) = ""
        For j = 1 To maxColumnSize
            dataRows(i, j) = ""
        Next
    Next
    
End Sub

'---- columnName operation
Function getColumnName(idx As Long) As String
   getColumnName = columeName(idx)
End Function

Sub setColumnName(idx As Long, name As String)
    columeName(idx) = name
    If (idx > cMax) Then
        cMax = idx
    End If
End Sub

'---- dataMap and fieldMap operation
Sub buildDataMap()
   Dim c As Long
   For c = 1 To cMax
        If (columeName(c) <> "") Then
            dataMapSize = dataMapSize + 1
            fieldMap(dataMapSize) = columeName(c)
            dataMap(dataMapSize) = c
        End If
   Next
End Sub

Function dataMapItem(idx As Long) As Long
    dataMapItem = dataMap(idx)
End Function

Function getFieldMap() As String()
   getFieldMap = fieldMap
End Function

Function fieldMapItem(idx As Long) As String
    fieldMapItem = fieldMap(idx)
End Function

Function getDataMapSize() As Long
    getDataMapSize = dataMapSize
End Function

'--- DataRows
Sub setDataItem(r As Long, c As Long, value As String)
    dataRows(r, c) = value
    If (dRowSize < r) Then
        dRowSize = r
    End If
End Sub

Function getDataItem(r As Long, c As Long) As String
    getDataItem = dataRows(r, c)
End Function

Function getDataItemSize() As Long
    getDataItemSize = dRowSize
End Function

Sub setSection(sec As String)
    section = sec
End Sub

Function getSection() As String
    getSection = section
End Function
