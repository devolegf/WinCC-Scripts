Sub ButtonConfiguration()
'VBA508
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateBorderColor = 1
    End With
End Sub
