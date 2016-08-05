Sub CreateAndPrintView()
'VBA177
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.PrintDocument
End Sub
