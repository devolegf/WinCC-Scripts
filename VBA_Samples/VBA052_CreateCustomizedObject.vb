Sub CreateCustomizedObject()
'VBA52
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.Selection.CreateCustomizedObject
    objCustomizedObject.ObjectName = "My Customized Object"
End Sub