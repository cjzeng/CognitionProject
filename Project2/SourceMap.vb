Option Explicit

Dim arraySize As Long
Dim columeName(1 To 400) As String

Function getColumnName(idx As Long) As String
   getColumnName = columeName(idx)
End Function

Sub setColumnName(idx As Long, name As String)
    columeName(idx) = name
End Sub

Function getColSize() As Long
    getColSize = arraySize
End Function

Sub setColSize(size As Long)
    arraySize = size
End Sub

