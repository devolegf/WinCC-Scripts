Sub CreateGroup()
'VBA43
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 80
        .Left = 80
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "myGroup"
End Sub