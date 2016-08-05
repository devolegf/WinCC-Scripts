Sub RemoveObjectFromGroup()
'VBA179
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
        .Width = 40
        .BackColor = RGB(255, 0, 0)
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    MsgBox "Group-object is created."
    objGroup.GroupedHMIObjects.Remove ("sEllipse")
    MsgBox "The ellipse is removed from group-object."
End Sub
