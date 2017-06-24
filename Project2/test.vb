Sub testSourceMap()
    Dim src As New SourceMap
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
