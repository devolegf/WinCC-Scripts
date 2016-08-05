Sub CreateGroup()
'VBA114
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
    'Set name for new group-object
    'The name identifies the group-object
    objGroup.ObjectName = "My Group"
    'Add new object to active document...
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    '...and add it to the group:
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub
