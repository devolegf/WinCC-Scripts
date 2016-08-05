Sub GraphicObjectConfiguration()
'VBA672
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicTransColor = 16711680
        .PicUseTransColor = True
    End With
End Sub
