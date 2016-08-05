Sub ApplySameWidthAndHeightToSelectedObjects()
'VBA183
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
    With objCircle
        .Top = 30
        .Left = 0
        .Height = 15
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 42
        .Width = 25
        .Height = 40
        .Selected = True
    End With
    With objEllipse
        .Top = 48
        .Left = 162
        .Width = 40
        .Height = 120
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    ActiveDocument.Selection.SameWidthAndHeight
End Sub
