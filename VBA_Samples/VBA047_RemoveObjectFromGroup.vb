Sub RemoveObjectFromGroup()
'VBA47
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    'delete group-object' first object
    objGroup.GroupedHMIObjects.Remove (1)
End Sub