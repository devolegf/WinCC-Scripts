Sub ShowCustomMenuInformation()
'VBA470
    Dim strKey As String
    Dim strLabel As String
    Dim strOutput As String
    Dim iIndex As Integer
    For iIndex = 1 To ActiveDocument.CustomMenus.Count
        strKey = ActiveDocument.CustomMenus(iIndex).Key
        strLabel = ActiveDocument.CustomMenus(iIndex).Label
        strOutput = strOutput & vbCrLf & "Key: " & strKey & "  Label: " & strLabel
    Next iIndex
    If 0 = ActiveDocument.CustomMenus.Count Then
        strOutput = "There are no custommenus for the document created."
    End If
    MsgBox strOutput
End Sub
