Sub ButtonConfiguration()
'VBA516
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTSIZE = 10
    End With
End Sub
