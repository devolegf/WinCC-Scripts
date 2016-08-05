Sub ShowGroupedObjectsOfFirstGroup()
'VBA265
    Dim colGroupedObjects As HMIGroupedObjects
    Dim objObject As HMIObject
    Set colGroupedObjects = ActiveDocument.HMIObjects("Group1").GroupedHMIObjects
    For Each objObject In colGroupedObjects
        MsgBox objObject.ObjectName
    Next objObject
End Sub
