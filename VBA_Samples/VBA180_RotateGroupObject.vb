Sub RotateGroupObject()
'VBA180
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    With objRectangle
        .Top = 30
        .Left = 30
        .Width = 80
        .Height = 40
        .Selected = True
    End With
    With objCircle
        .Top = 30
        .Left = 30
        .BackColor = RGB(255, 255, 255)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    MsgBox "Group-object created."
    objGroup.Selected = True
    ActiveDocument.Selection.Rotate
End Sub
