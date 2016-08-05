Sub ButtonConfiguration()
'VBA530
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Hotkey = 116
    End With
End Sub
