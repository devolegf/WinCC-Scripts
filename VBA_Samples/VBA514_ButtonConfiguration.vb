Sub ButtonConfiguration()
'VBA514
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTITALIC = True
    End With
End Sub
