Sub AddNewFolderToProjectLibrary()
'VBA20
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
'
    '("AddFolder(DefaultName)"-Methode):
    objProjectLib.FolderItems.AddFolder ("Custom Folder")
End Sub