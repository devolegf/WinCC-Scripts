Sub CreateCustomizedObject()
'VBA230
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Dim objCustomizedObject As HMICustomizedObject
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    With objCircle
        .Left = 10
        .Top = 10
        .Selected = True
    End With
    With objRectangle
        .Left = 50
        .Top = 50
        .Selected = True
    End With
    MsgBox "objects created and selected!"
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
    objCustomizedObject.ObjectName = "Customer-Object"
End Sub
