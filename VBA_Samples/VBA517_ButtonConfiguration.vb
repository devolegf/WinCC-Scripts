Sub ButtonConfiguration()
'VBA517
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FontUnderline = True
    End With
End Sub
