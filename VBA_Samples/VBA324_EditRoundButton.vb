Sub EditRoundButton()
'VBA324
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects("Roundbutton1")
    objRoundButton.BorderColor = RGB(255, 0, 0)
End Sub
