Sub ApplicationWindowConfig()
'VBA646
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .Moveable = True
    End With
End Sub
