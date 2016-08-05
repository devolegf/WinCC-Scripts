Private Sub Document_Activated(CancelForwarding As Boolean)
'VBA76
    MsgBox "The document got the focus." & vbCrLf & "This event (Document_Activated) is raised by the document itself"
End Sub