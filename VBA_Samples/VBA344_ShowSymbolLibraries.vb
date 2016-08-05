Sub ShowSymbolLibraries()
'VBA344
    Dim colSymbolLibraries As HMISymbolLibraries
    Dim objSymbolLibrary As HMISymbolLibrary
    Dim strLibraryList As String
    Set colSymbolLibraries = Application.SymbolLibraries
    For Each objSymbolLibrary In colSymbolLibraries
        strLibraryList = strLibraryList & objSymbolLibrary.Name & vbCrLf
    Next objSymbolLibrary
    MsgBox strLibraryList
End Sub
