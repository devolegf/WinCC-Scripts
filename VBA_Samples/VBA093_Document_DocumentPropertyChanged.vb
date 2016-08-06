Private Sub Document_DocumentPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA93
    Dim strPropName As String
    '"strPropName" contains the name of the modified property
    strPropName = Property.Name
    MsgBox "The picture-property " & strPropName & " is modified..."
End Sub