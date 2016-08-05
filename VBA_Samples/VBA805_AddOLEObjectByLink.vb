Sub AddOLEObjectByLink()
'VBA805
    Dim objOLEObject As HMIOLEObject
    Dim strFilename As String
'
    'Add OLEObject by filename. In this case, the filename has to
    'contain filename and path.
    'Replace the definition of strFilename with a filename with path
    'existing on your system
    strFilename = Application.ApplicationDataPath & "Test.bmp"
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("OLEObject1", strFilename, hmiOLEObjectCreationTypeByLink, False)
End Sub
