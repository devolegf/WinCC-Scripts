Sub ButtonConfiguration()
'VBA740
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .Text = "Button1"
    End With
End Sub
