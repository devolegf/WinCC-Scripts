Sub ChangePropertiesOfGroupMembers()
'VBA50
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objEllipse As HMIEllipse
    Dim objGroup As HMIGroup
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    Set objEllipse = ActiveDocument.HMIObjects.AddHMIObject("sEllipse", "HMIEllipse")
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
    With objEllipse
        .Top = 120
        .Left = 120
        .Selected = True
    End With
    MsgBox "Objects selected!"
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "My Group"
    'Set bordercolor of 1. object = "red":
    objGroup.GroupedHMIObjects(1).Properties("BorderColor") = RGB(255, 0, 0)
    'set x-coordinate of 2. object = "120" :
    objGroup.GroupedHMIObjects(2).Properties("Left") = 120
    'set y-coordinate of 3. object = "90":
    objGroup.GroupedHMIObjects(3).Properties("Top") = 90
End Sub
