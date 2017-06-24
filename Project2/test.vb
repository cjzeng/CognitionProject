Sub testSourceMap()
    Dim src As New SourceMap
    src.setColumnName 1, "first"
    src.setColumnName 2, "second"
    Debug.Print "idx 1=", src.getColumnName(1)
    Debug.Print "idx 2=", src.getColumnName(2)

End Sub

