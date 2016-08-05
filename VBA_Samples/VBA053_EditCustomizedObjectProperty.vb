Sub EditCustomizedObjectProperty()
'VBA53
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects(1)
    objCustomizedObject.Properties("BackColor") = RGB(255, 0, 0)
End Sub
