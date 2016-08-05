Sub ShowProjectInfo()
'VBA690
    Dim iProjectType As Integer
    Dim strProjectName As String
    Dim strProjectType As String
    iProjectType = Application.ProjectType
    strProjectName = Application.ProjectName
    Select Case iProjectType
        Case 0
            strProjectType = "Single-User System"
        Case 1
            strProjectType = "Multi-User System"
        Case 2
            strProjectType = "Client System"
    End Select
    MsgBox "Projecttype: " & strProjectType & vbCrLf & "Projectname: " & strProjectName
End Sub
