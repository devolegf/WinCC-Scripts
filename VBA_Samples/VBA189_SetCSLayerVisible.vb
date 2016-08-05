Sub SetCSLayerVisible()
'VBA189
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.SetCSLayerVisible 2, False
End Sub
