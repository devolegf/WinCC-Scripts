Sub ButtonConfiguration()
'VBA518
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .ForeColor = RGB(255, 0, 0)
    End With
End Sub
