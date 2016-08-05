Sub EditButton()
'VBA218
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects("Button")
    objButton.BorderColor = RGB(255, 0, 0)
End Sub
