Sub ApplicationWindowConfig()
'VBA797
    Dim objAppWindow As HMIApplicationWindow
    Set objAppWindow = ActiveDocument.HMIObjects.AddHMIObject("AppWindow", "HMIApplicationWindow")
    With objAppWindow
        .Caption = True
        .CloseButton = False
        .Height = 200
        .Left = 10
        .MaximizeButton = True
        .Moveable = False
        .OnTop = True
        .Sizeable = True
        .Top = 20
        .Visible = True
        .Width = 250
        .WindowBorder = True
    End With
End Sub
