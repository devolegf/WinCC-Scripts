Sub CopyObjectFromGlobalLibraryToProjectLibrary()
'VBA21
    Dim objGlobalLib As HMISymbolLibrary
    Dim objProjectLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'Copies object "PC" from the "Global Library" into the clipboard
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
'
    'The folder "Custom Folder" has to be available
    objProjectLib.FolderItems("Folder1").Folder.AddFromClipBoard ("Copy of PC/PLC")
End Sub
