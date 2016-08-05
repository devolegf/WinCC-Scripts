Sub ButtonConfiguration()
'VBA519
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeFlashColorOff = RGB(255, 0, 0)
    End With
End Sub
