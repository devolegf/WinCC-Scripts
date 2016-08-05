Sub ModifyPropertyOfObjectInGroup()
'VBA44
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("myGroup")
    objGroup.GroupedHMIObjects(1).Properties("BorderColor") = RGB(255, 0, 0)
End Sub
