Sub AddOLEObjectToActiveDocument()
'VBA40
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("MS Wordpad Document1", "Wordpad.Document.1")
End Sub
