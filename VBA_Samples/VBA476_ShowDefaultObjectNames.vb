Sub ShowDefaultObjectNames()
'VBA476
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To Application.DefaultHMIObjects.Count
        strOutput = strOutput & vbCrLf & Application.DefaultHMIObjects(iIndex).ObjectName
    Next iIndex
    MsgBox strOutput
End Sub
