Sub UnGroup()
'VBA48
    Dim objGroup As HMIGroup
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    objGroup.UnGroup
End Sub