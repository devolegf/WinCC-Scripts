Sub CreateGroup()
'VBA526
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipseSegment As HMIEllipseSegment
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
    objGroup.ObjectName = "Group1"
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
'
    'Add one object to the existing group
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub
