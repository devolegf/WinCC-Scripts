Sub DestroyCustomizedObject()
'VBA55
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("My Customized Object")
    objCustomizedObject.Destroy
End Sub