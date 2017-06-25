Option Explicit

'Dim maxArraySize As Long

Const maxArraySize = 400
Const maxColumnSize = 30

Const fnIdx = 1
Const lnIdx = 2
Const secIdx = 3
Const dataStart = 3

'columnName - maps index to header
Dim cMax As Long  'columnName actaul size used
Dim oColumnSize As Long 'previous column size
Dim columnName(1 To maxArraySize) As String 'columnName array

'dataMap - maps of indexs to usable data column
Dim dataMap(1 To maxArraySize) As Long
Dim dataMapSize As Long 'The length of dataMap array

'fieldMap - field name for each index in data map, uses dataMapSize for actual size
'Dim fieldMap(1 To maxArraySize) As String

'Dim dataRows(1 To maxArraySize, 1 To maxColumnSize) As String
'Dim dRowSize As Long

Sub init()
    cMax = 1
    oColumnSize = 1
    dataMapSize = 0
'   Dim i As Long
'  Dim j As Long
'    For i = 1 To maxArraySize
'        columeName(i) = ""
'        dataMap(i) = 0
'        fieldMap(i) = ""
'        For j = 1 To maxColumnSize
'            dataRows(i, j) = ""
'        Next
'    Next
    
End Sub

'---- columnName operation
Sub buildColumnMap(ByRef fieldMap() As String, size As Long)
    oColumnSize = cMax
    Dim i As Long
    ' Build initial one
    If cMax = 1 Then
       columnName(fnIdx) = fieldMap(fnIdx)
       columnName(lnIdx) = fieldMap(lnIdx)
       columnName(secIdx) = "Section"
       dataMap(fnIdx) = fnIdx
       dataMap(lnIdx) = lnIdx
       cMax = 3
    End If
    For i = dataStart To size
        setColumnMapItem i, fieldMap(i)
    Next
End Sub

Sub setColumnMapItem(idx As Long, fieldName As String)
    Dim foundIdx As Long
    foundIdx = findFieldName(fieldName)
    If (foundIdx = 0) Then
        cMax = cMax + 1
        columnName(cMax) = fieldName
        foundIdx = cMax
    End If
    dataMap(idx) = foundIdx
End Sub

Function getDataMapIdx(idx As Long) As Long
    getDataMapIdx = dataMap(idx)
End Function

Function getColumnMapItem(idx As Long) As String
    getColumnMapItem = columnName(idx)
End Function


Function findFieldName(val As String)
    Dim i As Long
    Dim found As Long
    found = 0
    For i = 1 To cMax
        If (columnName(i) = val) Then
            found = i
            Exit For
        End If
    Next
    findFieldName = found
End Function

Function getColumnName(idx As Long) As String
   getColumnName = columeName(idx)
End Function

Sub setColumnName(idx As Long, name As String)
    columeName(idx) = name
    If (idx > cMax) Then
        cMax = idx
    End If
End Sub
 
Function getCMapOIdx() As Long
    getCMapOIdx = oColumnSize
End Function

Function getCMapSize() As Long
    getCMapSize = cMax
End Function
