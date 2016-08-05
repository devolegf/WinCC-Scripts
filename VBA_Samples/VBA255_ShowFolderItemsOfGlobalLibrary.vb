Sub ShowFolderItemsOfGlobalLibrary()
'VBA255
    Dim colFolderItems As HMIFolderItems
    Dim objFolderItem As HMIFolderItem
    Set colFolderItems = Application.SymbolLibraries(1).FolderItems
    For Each objFolderItem In colFolderItems
        MsgBox objFolderItem.Name
    Next objFolderItem
End Sub
