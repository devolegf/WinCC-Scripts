Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA539
    Dim objProp As HMIProperty
    Dim strStatus As String
    Set objProp = Property
'
    'Checks whether property is dynamicable
    If objProp.IsDynamicable = True Then
        strStatus = "yes"
    Else
        strStatus = "no"
    End If
    MsgBox "Property: " & objProp.Name & vbCrLf & "Value: " & objProp.value & vbCrLf & "Dynamicable: " & strStatus
End Sub
