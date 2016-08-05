Sub ButtonConfiguration()
'VBA507
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashRateBackColor = 1
    End With
End Sub
