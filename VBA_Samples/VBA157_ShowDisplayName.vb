Sub ShowDisplayName()
'VBA157
    Dim objGlobalLib As HMISymbolLibrary
    Dim objFItem As HMIFolderItem
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objFItem = objGlobalLib.GetItemByPath("\Folder1\Folder2\Object1")
    MsgBox objFItem.DisplayName
End Sub
