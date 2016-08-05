Sub ApplicationWindowConfig()
'VBA619
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .MaximizeButton = True
    End With
End Sub
