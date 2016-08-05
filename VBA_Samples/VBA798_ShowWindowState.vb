Sub ShowWindowState()
'VBA798
    Dim strState As String
    Select Case Application.WindowState
        Case 0
            strState = "The application-window is maximized"
        Case 1
            strState = "The applicationwindow is minimized"
        Case 2
            strState = "The application-window has a userdefined size"
    End Select
    MsgBox strState
End Sub
