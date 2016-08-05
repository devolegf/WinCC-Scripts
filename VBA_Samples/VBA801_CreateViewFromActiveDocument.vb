Sub CreateViewFromActiveDocument()
'VBA801
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Zoom = 50
End Sub
