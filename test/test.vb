Sub testSourceMap()
    Dim src As New SourceMap
    src.init
    src.setColumnName 1, "first"
    src.setColumnName 2, "second"
    Debug.Print "idx 1=", src.getColumnName(1)
    Debug.Print "idx 2=", src.getColumnName(2)

End Sub

Sub testNameLookUp()
    Dim lookUp As New NameLookUp
    
    lookUp.init
    
    lookUp.add "name 1"
    lookUp.add "name 2"
    lookUp.add "name 3"
    lookUp.add "name 4"
    lookUp.add "name 5"
    lookUp.add "name 6"
    lookUp.add "name 7"
    
    Debug.Print "name 6 = ", lookUp.findNameIdx("name 6")
End Sub

Sub testDataMap()
    Dim src As New SourceMap
    src.init
    src.setColumnName 1, "first"
    src.setColumnName 2, "second"
    src.setColumnName 4, "third"
    src.setColumnName 8, "four"
    src.setColumnName 10, "five"
    
    src.buildDataMap
    
    Dim i As Long
    For i = 1 To src.getDataMapSize
        Debug.Print "i=", i, ", item=", src.fieldMapItem(i), ", fieldIdx=", src.dataMapItem(i)
        
    Next
End Sub

Sub testDataRows()
    Dim src As New SourceMap
    src.init
    src.setColumnName 1, "first"
    src.setColumnName 2, "second"
    src.setColumnName 4, "third"
    src.setColumnName 8, "four"
    src.setColumnName 10, "five"
    
    src.buildDataMap
    
    Dim c As Long
    Dim r As Long
    For r = 1 To 10
        For c = 1 To src.getDataMapSize
           src.setDataItem r, c, "item" & r & c
        Next
    Next
    For r = 1 To 10
        For c = 1 To src.getDataMapSize
           Debug.Print "Item:", r, ":", c, " is ", src.getDataItem(r, c)
        Next
    Next


End Sub

Sub testDestMap()
    Dim src As New SourceMap
    src.init
    src.setColumnName 1, "firstName"
    src.setColumnName 2, "lastName"
    src.setColumnName 4, "third"
    src.setColumnName 8, "four"
    src.setColumnName 10, "five"
    
    src.buildDataMap
    
    Dim Dest As New Dest
    Dim oIdx As Long
    Dim i As Long
    Dim dataLength As Long

    Dest.init
    dataLength = src.getDataMapSize
    Dest.buildColumnMap src.getFieldMap, dataLength
    
    oIdx = Dest.getCMapOIdx()
    For i = oIdx To Dest.getCMapSize()
        Debug.Print "colmap ", i, "=", Dest.getColumnMapItem(i)
    Next
    
    For i = 1 To dataLength
        Debug.Print "dataMap ", i, "=", Dest.getDataMapIdx(i)
    Next
    
    src.init
    src.setColumnName 1, "firstName"
    src.setColumnName 2, "lastName"
    src.setColumnName 4, "third"
    src.setColumnName 8, "six"
    src.setColumnName 10, "seven"
    
    src.buildDataMap
    
    dataLength = src.getDataMapSize
    Dest.buildColumnMap src.getFieldMap, dataLength
     
    For i = 1 To Dest.getCMapSize()
        Debug.Print "colmap ", i, "=", Dest.getColumnMapItem(i)
    Next
    
    For i = 1 To dataLength
        Debug.Print "dataMap ", i, "=", Dest.getDataMapIdx(i)
    Next
    
End Sub

