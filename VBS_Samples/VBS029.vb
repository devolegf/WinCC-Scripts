Sub OnClick(ByVal Item)
'VBS29
    Dim objScreenItem
    On Error Resume Next    'Activation of errorhandling
    For Each objScreenItem In ScreenItems
        If "HMIRectangle" = objScreenItem.Type Then
'
            '=== Property "RoundCornerHeight" available only for RoundRectangle
            objScreenItem.RoundCornerHeight = objScreenItem.RoundCornerHeight * 2
            If 0 <> Err.Number Then
                HMIRuntime.Trace objScreenItem.ObjectName & ": no RoundedRectangle" & vbCrLf
                Err.Clear    'Delete errormessage
            End If
        End If
    Next
    On Error Goto 0    'Deactivation of errorhandling
End Sub