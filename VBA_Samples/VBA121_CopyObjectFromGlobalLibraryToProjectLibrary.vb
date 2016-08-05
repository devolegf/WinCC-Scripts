Sub CopyObjectFromGlobalLibraryToProjectLibrary()
'VBA121
    Dim objGlobalLib As HMISymbolLibrary
    Dim objProjectLib As HMISymbolLibrary
    Set objGlobalLib = Application.SymbolLibraries(1)
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder3")
'
    'copy object from "Global Library" to clipboard
    With objGlobalLib
        .FolderItems(2).Folder.Item(2).Folder.Item(1).CopyToClipboard
    End With
'
    'paste object from clipboard into "Project Library"
    objProjectLib.FolderItems(objProjectLib.FindByDisplayName("My Folder3")).Folder.AddFromClipBoard ("Copy of PC/PLC")
End Sub
