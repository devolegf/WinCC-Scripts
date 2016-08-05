Sub CopyFolderItemToClipboard()
'VBA254
    Dim objGlobalLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
End Sub
