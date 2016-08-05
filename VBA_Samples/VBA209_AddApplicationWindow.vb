Sub AddApplicationWindow()
'VBA209
    Dim objApplicationWindow As HMIApplicationWindow
    Set objApplicationWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow", "HMIApplicationWindow")
End Sub
