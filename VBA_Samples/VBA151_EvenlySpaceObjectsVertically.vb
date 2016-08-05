Sub EvenlySpaceObjectsVertically()
'VBA151
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects created and selected"
    ActiveDocument.Selection.EvenlySpaceVertically
End Sub
