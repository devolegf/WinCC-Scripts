Sub CreateViewAndActivateView()
'VBA712
    Dim objView As HMIView
    Set objView = ActiveDocument.Views.Add
    objView.Activate
    objView.ScrollPosX = 40
    objView.ScrollPosY = 10
End Sub
