Sub ShowFirstObjectOfCollection()
'VBA268
    Dim strName As String
    strName = ActiveDocument.HMIObjects(1).ObjectName
    MsgBox strName
End Sub
