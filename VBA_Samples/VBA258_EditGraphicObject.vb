Sub EditGraphicObject()
'VBA258
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects("Graphic-Object")
    objGraphicObject.BorderColor = RGB(255, 0, 0)
End Sub
