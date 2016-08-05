Private Sub Document_HMIObjectAdded(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA94
    Dim strObjName As String
'
    '"strObjName" contains the name of the added object
    strObjName = HMIObject.ObjectName
    MsgBox "Object " & strObjName & " is added..."
End Sub