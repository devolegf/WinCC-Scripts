Sub IOFieldConfiguration()
'VBA372
    Dim objIOField As HMIIOField
'
    'Add new IO-Feld to active document:
    Set objIOField = ActiveDocument.HMIObjects.AddHMIObject("IOField1", "HMIIOField")
    With objIOField
        .AdaptBorder = True
    End With
End Sub
