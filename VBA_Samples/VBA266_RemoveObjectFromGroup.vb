Sub RemoveObjectFromGroup()
'VBA266
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("Group1")
    objGroup.GroupedHMIObjects.Remove (1)
End Sub
