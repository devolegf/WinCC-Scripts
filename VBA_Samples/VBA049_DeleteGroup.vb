Sub DeleteGroup()
'VBA49
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    objGroup.Delete
End Sub