Private Sub Document_HMIObjectMoved(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA95
    Dim strObjName As String
'
    '"strObjName" contains the name of the moved object
    strObjName = HMIObject.ObjectName
    MsgBox "Object " & strObjName & " was moved..."
End Sub