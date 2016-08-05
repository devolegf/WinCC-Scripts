Sub ExampleForParent()
'VBA661
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    MsgBox objView.Parent.Name
End Sub
