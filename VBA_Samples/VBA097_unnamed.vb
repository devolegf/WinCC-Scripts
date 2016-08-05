Private Sub Document_HMIObjectResized(ByVal HMIObject As IHMIObject, CancelForwarding As Boolean)
'VBA97
    Dim strObjName As String
'
    '"strObjName" contains the name of the modified object
    strObjName = HMIObject.ObjectName
    MsgBox "The size of " & strObjName & " was modified..."
End Sub