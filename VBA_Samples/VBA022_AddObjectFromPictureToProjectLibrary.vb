Sub AddObjectFromPictureToProjectLibrary()
'VBA22
    Dim objProjectLib As HMISymbolLibrary
    Dim objCircle As HMICircle
    Set objProjectLib = Application.SymbolLibraries(2)
'
    'Insert new object "Circle1"
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
'
    'The folder "Custom Folder" has to be available
    '("AddItem(DefaultName, pHMIObject)"-Methode):
    objProjectLib.FolderItems("Folder1").Folder.AddItem "ProjectLib Circle", ActiveDocument.HMIObjects("Circle1")
End Sub