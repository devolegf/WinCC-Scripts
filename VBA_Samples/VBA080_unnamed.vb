Private Sub Document_BeforeHMIObjectDelete(ByVal HMIObject As IHMIObject, Cancel As Boolean, CancelForwarding As Boolean)
'VBA80
    Dim strObjName As String
    Dim strAnswer As String
'
    '"strObjName" contains the name of the deleted object
    strObjName = HMIObject.ObjectName
    strAnswer = MsgBox("Are you sure to delete " & strObjName & "?", vbYesNo)
    If strAnswer = vbNo Then
        'if pressed "No" -> set Cancel to true for prevent delete
        Cancel = True
    End If
End Sub