Sub CopyObjectFromGlobalLibraryToActiveDocument()
'VBA24
    Dim objGlobalLib As HMISymbolLibrary
    Dim objHMIObject As HMIObject
    Dim iLastObject As Integer
    Set objGlobalLib = Application.SymbolLibraries(1)
'
    'Copy object "PC" from "Global Library" to clipboard
    objGlobalLib.FolderItems("Folder2").Folder("Folder2").Folder.Item("Object1").CopyToClipboard
'
    'Get object from clipboard and add it to active document
    ActiveDocument.PasteClipboard
'
    'Get last inserted object
    iLastObject = ActiveDocument.HMIObjects.Count
    Set objHMIObject = ActiveDocument.HMIObjects(iLastObject)
'
    'Set position of the object:
    With objHMIObject
        .Left = 40
        .Top = 40
    End With
End Sub