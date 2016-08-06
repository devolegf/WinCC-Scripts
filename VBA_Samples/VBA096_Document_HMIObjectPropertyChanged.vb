Private Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA96
    Dim strObjProp As String
    Dim strObjName As String
    Dim varPropValue As Variant
'
    '"strObjProp" contains the name of the modified property
    '"varPropValue" contains the new value
    strObjProp = Property.Name
    varPropValue = Property.value
'
    '"strObjName" contains the name of the selected object,
    'which property is modified
    strObjName = Property.Application.ActiveDocument.Selection(1).ObjectName
    MsgBox "The property " & strObjProp & " of object " & strObjName & " is modified... " & vbCrLf & "The new value is: " & varPropValue
End Sub
