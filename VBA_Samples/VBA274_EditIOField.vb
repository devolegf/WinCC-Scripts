Sub EditIOField()
'VBA274
    Dim objIOField As HMIIOField
    Set objIOField = ActiveDocument.HMIObjects("IO-Field")
    objIOField.BorderColor = RGB(255, 0, 0)
End Sub
