Sub ShowSelectionOfDocument()
'VBA331
    Dim colSelection As HMISelectedObjects
    Dim objObject As HMIObject
    Dim strObjectList As String
    Set colSelection = ActiveDocument.Selection
    If colSelection.Count <> 0 Then
        strObjectList = "List of selected objects:"
        For Each objObject In colSelection
            strObjectList = strObjectList & vbCrLf & objObject.ObjectName
        Next objObject
    Else
        strObjectList = "No objects selected"
    End If
    MsgBox strObjectList
End Sub
