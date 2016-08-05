Sub ShowCustomToolbarInformation()
'VBA471
    Dim strKey As String
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomToolbars.Count
        strKey = ActiveDocument.CustomToolbars(iIndex).Key
        strOutput = strOutput & vbCrLf & "Key: " & strKey
    Next iIndex
    If 0 = ActiveDocument.CustomToolbars.Count Then
        strOutput = "There are no toolbars created for this document."
    End If
    MsgBox strOutput
End Sub
