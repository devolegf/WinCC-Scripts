Sub AddObjectToGroup()
'VBA46
    Dim objGroup As HMIGroup
    Dim objEllipseSegment As HMIEllipseSegment
    'Adds new object to active document
    Set objEllipseSegment = ActiveDocument.HMIObjects.AddHMIObject("EllipseSegment", "HMIEllipseSegment")
    Set objGroup = ActiveDocument.HMIObjects("My Group")
    'Adds the object to the group
    objGroup.GroupedHMIObjects.Add ("EllipseSegment")
End Sub