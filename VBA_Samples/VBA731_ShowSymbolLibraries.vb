Sub ShowSymbolLibraries()
'VBA731
    Dim colSymbolLibraries As HMISymbolLibraries
    Dim objSymbolLibrary As HMISymbolLibrary
    Set colSymbolLibraries = Application.SymbolLibraries
    For Each objSymbolLibrary In colSymbolLibraries
        MsgBox objSymbolLibrary.Name
    Next objSymbolLibrary
End Sub
