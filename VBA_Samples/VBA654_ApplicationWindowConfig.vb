Sub ApplicationWindowConfig()
'VBA654
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow1", "HMIApplicationWindow")
    With objAppWindow
        .OnTop = True
    End With
End Sub
