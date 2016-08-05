Sub EditApplicationWindow()
'VBA210
    Dim objApplicationWindow As HMIApplicationWindow
    Set objApplicationWindow = ActiveDocument.HMIObjects("AppWindow")
    objApplicationWindow.Sizeable = True
End Sub
