Sub CheckModificationOfActiveDocument()
'VBA645
    Dim strCheck As String
    Dim bModified As Boolean
    bModified = ActiveDocument.Modified
    Select Case bModified
        Case True
            strCheck = "Active document is modified"
        Case False
            strCheck = "Active document is not modified"
    End Select
    MsgBox strCheck
End Sub
