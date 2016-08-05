Sub DeleteCustomizedObject()
'VBA56
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("My Customized Object")
    objCustomizedObject.Delete
End Sub