Sub EditOLEObject()
'VBA299
    Dim objOleObject As HMIOLEObject
    Set objOleObject = ActiveDocument.HMIObjects("Wordpad Document")
    objOleObject.Left = 140
End Sub
