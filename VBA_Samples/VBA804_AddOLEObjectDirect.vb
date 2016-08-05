Sub AddOLEObjectDirect()
'VBA804
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("OLEObject1", "Wordpad.Document.1", hmiOLEObjectCreationTypeDirect, True)
End Sub
