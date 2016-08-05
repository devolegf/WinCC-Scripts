Sub ApplicationWindowConfig()
'VBA414
    Dim objAppWindow As HMIApplicationWindow
'
    'Add new applicationwindow to active document:
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .WindowBorder = True
    End With
End Sub
