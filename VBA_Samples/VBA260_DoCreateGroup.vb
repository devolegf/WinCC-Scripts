Sub DoCreateGroup()
'VBA260
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.Selection.CreateGroup
    objGroup.ObjectName = "Group-Object"
End Sub
