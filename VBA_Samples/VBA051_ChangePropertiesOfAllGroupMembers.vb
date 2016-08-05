Sub ChangePropertiesOfAllGroupMembers()
'VBA51
    Dim objGroup As HMIGroup
    Dim iMaxMembers As Integer
    Dim iIndex As Integer
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    iIndex = 1
'
    'Get number of objects in group-object:
    iMaxMembers = objGroup.GroupedHMIObjects.Count
'
    'set linecolor of all objects = "yellow":
    For iIndex = 1 To iMaxMembers
        objGroup.GroupedHMIObjects(iIndex).Properties("BorderColor") = RGB(255, 255, 0)
    Next iIndex
End Sub
