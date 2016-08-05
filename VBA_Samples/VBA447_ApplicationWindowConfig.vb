Sub ApplicationWindowConfig()
'VBA447
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .CloseButton = True
    End With
End Sub
