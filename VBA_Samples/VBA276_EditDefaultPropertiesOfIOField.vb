Sub EditDefaultPropertiesOfIOField()
'VBA276
    Dim objIOField As HMIIOField
    Set objIOField = Application.DefaultHMIObjects("HMIIOField")
    objIOField.BorderColor = RGB(255, 255, 0)
End Sub
