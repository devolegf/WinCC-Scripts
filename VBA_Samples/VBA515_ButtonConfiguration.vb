Sub ButtonConfiguration()
'VBA515
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTNAME = "Arial"
    End With
End Sub
