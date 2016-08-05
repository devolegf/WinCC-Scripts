Sub EditCustomizedObject()
'VBA231
    Dim objCustomizedObject As HMICustomizedObject
    Set objCustomizedObject = ActiveDocument.HMIObjects("Customer-Object")
    MsgBox objCustomizedObject.ObjectName
End Sub
