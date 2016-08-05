Sub ButtonConfiguration()
'VBA520
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeFlashColorOff = RGB(255, 255, 255)
    End With
End Sub
