Sub ShowObjectsOfDocument()
'VBA270
    Dim colObjects As HMIObjects
    Dim objObject As HMIObject
    Set colObjects = ActiveDocument.HMIObjects
    For Each objObject In colObjects
        MsgBox objObject.ObjectName
    Next objObject
End Sub
