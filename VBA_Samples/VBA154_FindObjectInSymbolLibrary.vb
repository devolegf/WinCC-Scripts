Sub FindObjectInSymbolLibrary()
'VBA154
    Dim objGlobalLib As HMISymbolLibrary
    Dim objFItem As HMIFolderItem
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objFItem = objGlobalLib.FindByDisplayName("PC")
    MsgBox objFItem.DisplayName
End Sub
