Sub EditGroup()
'VBA261
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("Group-Object")
    MsgBox objGroup.ObjectName
End Sub
