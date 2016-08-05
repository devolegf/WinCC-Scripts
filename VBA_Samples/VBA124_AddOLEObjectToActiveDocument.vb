Sub AddOLEObjectToActiveDocument()
'VBA124
    Dim objOLEObject As HMIOLEObject
    Set objOLEObject = ActiveDocument.HMIObjects.AddOLEObject("MS Wordpad Document", "Wordpad.Document.1")
End Sub
