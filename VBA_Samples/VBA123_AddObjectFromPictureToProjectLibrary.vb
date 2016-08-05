Sub AddObjectFromPictureToProjectLibrary()
'VBA123
    Dim objProjectLib As HMISymbolLibrary
    Dim objCircle As HMICircle
 
    Set objProjectLib = Application.SymbolLibraries(2)
    objProjectLib.FolderItems.AddFolder ("My Folder2")
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
'
    'Add object "Circle" to "Project Library":
    objProjectLib.FolderItems(objProjectLib.FindByDisplayName("My Folder2")).Folder.AddItem "ProjectLib Circle", ActiveDocument.HMIObjects("Circle")
End Sub
