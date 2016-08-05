Sub GraphicObjectConfiguration()
'VBA671
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("GraphicObject1", "HMIGraphicObject")
    With objGraphicObject
        .PicReferenced = True
    End With
End Sub
