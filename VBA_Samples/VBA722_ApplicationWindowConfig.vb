Sub ApplicationWindowConfig()
'VBA722
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .Sizeable = True
    End With
End Sub
