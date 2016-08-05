Sub CreateCustomizedObject()
'VBA54
    Dim objCustomizedObject As HMICustomizedObject
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
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
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
'
    '*** The "Configurationdialog" started. ***
    '*** Configure the costumize-object with the "configurationdialog" ***
'
    objCustomizedObject.ObjectName = "My Customized Object"
End Sub