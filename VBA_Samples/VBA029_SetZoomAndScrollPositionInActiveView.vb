Sub SetZoomAndScrollPositionInActiveView()
'VBA29
    Dim objView As HMIView 
    Set objView = ActiveDocument.Views.Add 
    With objView 
        .Activate 
        .ScrollPosX = 40 
        .ScrollPosY = 10 
        .Zoom = 150
    End With
End Sub