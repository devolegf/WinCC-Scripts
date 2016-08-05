Sub LayerInfo()
'VBA588
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        With objLayer
            iAnswer = MsgBox("Layername: " & .Name & vbCrLf & "max. zoom:  " & .MaxZoom & vbCrLf & "min. zoom:  " & .MinZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub
