Sub ConnectCheck()
'VBA538
    Dim bCheck As Boolean
    Dim strStatus As String
    bCheck = Application.IsConnectedToProject
    If bCheck = True Then
        strStatus = "yes"
    Else
        strStatus = "no"
    End If
    MsgBox "Connection to project available: " & strStatus
End Sub
