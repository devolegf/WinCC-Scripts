Sub AddGraphicObject()
'VBA257
    Dim objGraphicObject As HMIGraphicObject
    Set objGraphicObject = ActiveDocument.HMIObjects.AddHMIObject("Graphic-Object", "HMIGraphicObject")
End Sub
