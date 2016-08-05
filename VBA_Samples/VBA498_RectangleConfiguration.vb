Sub RectangleConfiguration()
'VBA498
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        .FlashBackColor = True
    End With
End Sub
