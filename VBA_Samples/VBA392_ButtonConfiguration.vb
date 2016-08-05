Sub ButtonConfiguration()
'VBA392
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .BackBorderWidth = 2
    End With
End Sub
