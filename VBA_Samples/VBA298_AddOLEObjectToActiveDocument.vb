Sub AddOLEObjectToActiveDocument()
'VBA298
    Dim objOleObject As HMIOLEObject
    Set objOleObject = ActiveDocument.HMIObjects.AddOLEObject("Wordpad Document", "Wordpad.Document.1")
End Sub
