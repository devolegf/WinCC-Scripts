Sub ButtonConfiguration()
'VBA510
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateForeColor = 1
    End With
End Sub
