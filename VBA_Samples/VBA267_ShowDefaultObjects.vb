Sub ShowDefaultObjects()
'VBA267
    Dim strType As String
    Dim strName As String
    Dim strMessage As String
    Dim iMax As Integer
    Dim iIndex As Integer
 
    iMax = Application.DefaultHMIObjects.Count
    iIndex = 1
    For iIndex = 1 To iMax
        With Application.DefaultHMIObjects(iIndex)
            strType = .Type
            strName = .ObjectName
            strMessage = strMessage & "Element: " & iIndex & " / Objecttype: " & strType & " / Objectname: " & strName
        End With
        If 0 = iIndex Mod 10 Then
            MsgBox strMessage
            strMessage = ""
        Else
            strMessage = strMessage & vbCrLf & vbCrLf
        End If
    Next iIndex
    MsgBox "Element: " & iIndex & vbCrLf & "Objecttype: " & strType & vbCrLf & "Objectname: " & strName
End Sub
