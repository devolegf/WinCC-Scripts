Sub ButtonConfiguration()
'VBA513
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .FONTBOLD = True
    End With
End Sub
