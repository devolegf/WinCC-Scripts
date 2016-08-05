Sub GraphicObjectConfiguration()
'VBA680
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicTransColor = RGB(0, 0, 255)
        .PicUseTransColor = True
    End With
End Sub
