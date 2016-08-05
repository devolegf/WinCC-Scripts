Sub Document_HMIObjectPropertyChanged(ByVal Property As IHMIProperty, CancelForwarding As Boolean)
'VBA72
    CancelForwarding = True
    MsgBox "Object's property has been changed!"
End Sub