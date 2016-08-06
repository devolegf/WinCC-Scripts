Private Sub Document_ViewCreated(ByVal pView As IHMIView, CancelForwarding As Boolean)
'VBA109
    Dim iViewCount As Integer
'
    'To read out the number of views
    iViewCount = pView.Application.ActiveDocument.Views.Count
    MsgBox "A new copy of the picture (number " & iViewCount & ") was created."
End Sub
