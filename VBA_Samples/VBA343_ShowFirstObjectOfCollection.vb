Sub ShowFirstObjectOfCollection()
'VBA343
    Dim strName As String
    strName = Application.SymbolLibraries(1).Name
    MsgBox strName
End Sub
