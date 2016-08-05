Sub ShowLayerWithNumbers()
'VBA803
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim iAnswer As Integer
    Dim iIndex As Integer
    iIndex = 1
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        iAnswer = MsgBox("Layername: " & objLayer & vbCrLf & "Layernumber: " & objLayer.Number & vbCrLf & "Layersindex: " & iIndex, vbOKCancel)
        iIndex = iIndex + 1
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub
