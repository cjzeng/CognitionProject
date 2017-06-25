' Class NameLookUp

Dim names(1 To 1000) As String
Dim cur As Long

Sub init()
    cur = 0
End Sub

Sub add(key As String)
    Dim i As Long
    i = cur
    cur = i + 1
    Debug.Print "cur=" & cur & ", key=" & key
    
    names(cur) = key
End Sub

Function findNameIdx(key As String) As Long
    Dim foundIdx As Long
    Dim idx As Long
    
    foundIdx = 0
    For idx = 1 To cur
        If key = names(idx) Then
            foundIdx = idx
            Exit For
        End If
    Next
    findNameIdx = foundIdx
End Function

Function getNameIdx(key As String) As Long
    Dim idx As Long
    idx = findNameIdx(key)
    If (idx = 0) Then
        add (key)
        idx = cur
    End If
    getNameIdx = idx
End Function
