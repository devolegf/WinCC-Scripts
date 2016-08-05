Sub ButtonConfiguration()
'VBA658
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Width = 150
        .Height = 150
        .Text = "Text is displayed vertical"
        .Orientation = False
    End With
End Sub
