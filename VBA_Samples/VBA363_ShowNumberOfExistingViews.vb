Sub ShowNumberOfExistingViews()
'VBA363
    Dim iMaxViews As Integer
    iMaxViews = ActiveDocument.Views.Count
    MsgBox "Number of copies from active document: " & iMaxViews
End Sub
