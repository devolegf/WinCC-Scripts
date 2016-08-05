Sub ShowInternalNameOfFolderItem()
'VBA536
    Dim objGlobalLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    MsgBox objGlobalLib.FolderItems(2).Folder(2).Folder.Item(1).Name
End Sub
