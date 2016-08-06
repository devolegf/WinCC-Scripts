Sub OnClick(ByVal Item)
'VBS27
    Dim objScreenItem
'
    'Activation of errorhandling:
    On Error Resume Next
    For Each objScreenItem In ScreenItems
        If "HMIRectangle" = objScreenItem.Type Then
'
            '=== Property "RoundCornerHeight" only available for RoundRectangle
            objScreenItem.RoundCornerHeight = objScreenItem.RoundCornerHeight * 2
            If 0 <> Err.Number Then
                HMIRuntime.Trace objScreenItem.Name & ": no RoundedRectangle" & vbCrLf
'
                'Delete error message
                Err.Clear
            End If
        End If
    Next
    On Error Goto 0  'Deactivation of errorhandling
End Sub