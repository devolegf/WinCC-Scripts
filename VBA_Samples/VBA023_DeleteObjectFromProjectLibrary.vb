Sub DeleteObjectFromProjectLibrary()
'VBA23
    Dim objProjectLib As HMISymbolLibrary
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'The folder "Custom Folder" has to be available
    '("Delete"-Methode):
    objProjectLib.FolderItems("Folder1").Delete
End Sub