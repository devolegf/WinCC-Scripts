Sub ShowPropertiesDynamicsofAllObjects()
'VBA483
    Dim objObject As HMIObject
    Dim colObjects As HMIObjects
    Dim colProperties As HMIProperties
    Dim objProperty As HMIProperty
    Dim strOutput As String
    Set colObjects = Application.ActiveDocument.HMIObjects
    For Each objObject In colObjects
        Set colProperties = objObject.Properties
        For Each objProperty In colProperties
            If 0 <> objProperty.DynamicStateType Then
                strOutput = strOutput & vbCrLf & objObject.ObjectName & " - " & objProperty.DisplayName & ": Statetype " & objProperty.Dynamic.DynamicStateType
            End If
        Next objProperty
    Next objObject
    MsgBox strOutput
End Sub
