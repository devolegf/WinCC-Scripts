Sub ButtonConfiguration()
'VBA501
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FlashForeColor = True
    End With
End Sub
