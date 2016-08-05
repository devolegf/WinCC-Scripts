Sub ButtonConfiguration()
'VBA418
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
        .BorderColorBottom = RGB(255, 0, 0)
        .BorderColorTop = RGB(0, 0, 255)
    End With
End Sub
