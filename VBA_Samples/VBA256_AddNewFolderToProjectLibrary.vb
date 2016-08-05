Sub AddNewFolderToProjectLibrary()
'VBA256
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder")
End Sub
