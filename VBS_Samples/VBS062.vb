VBS62
    Dim objScreenItem
    On Error Resume Next    'Activation of errorhandling
    For Each objScreenItem In ScreenItems
        If objScreenItem.Type = "HMIButton" Then
'
            '=== Property "Text" available only for Standard-Button
            objScreenItem.Text = "Windows"
            If 0 <> Err.Number Then
                HMIRuntime.Trace objScreenItem.ObjectName & ": no Windows-Button" & vbCrLf
                Err.Clear    'Delete error message
            End If
'
            '=== Property "Radius" available only for RoundButton
            objScreenItem.Radius = 10
            If 0 <> Err.Number Then
                HMIRuntime.Trace objScreenItem.ObjectName & ": no RoundButton" & vbCrLf
                Err.Clear
            End If
'
            '--- Property "Caption" available only for PushButton
            objScreenItem.Caption = "Push"
            If 0 <> Err.Number Then
                HMIRuntime.Trace objScreenItem.ObjectName & ": no Control" & vbCrLf
                Err.Clear
            End If
        End If
    Next
    On Error Goto 0    'Deactivation of errorhandling