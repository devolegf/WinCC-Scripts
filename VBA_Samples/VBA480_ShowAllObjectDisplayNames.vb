Sub ShowAllObjectDisplayNames()
'VBA480
    Dim strOutput As String
    Dim iIndex1 As Integer
    iIndex1 = 1
    strOutput = "List of all properties-displaynames from object """ & Application.DefaultHMIObjects(1).ObjectName & """" & vbCrLf & vbCrLf
    For iIndex1 = 1 To Application.DefaultHMIObjects(1).Properties.Count
        strOutput = strOutput & Application.DefaultHMIObjects(1).Properties(iIndex1).DisplayName & " / "
    Next iIndex1
    MsgBox strOutput
End Sub
