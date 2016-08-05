Sub LayerInfo()
'VBA644
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim strMaxZoom As String
    Dim strMinZoom As String
    Dim strLayerName As String
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        With objLayer
            strMinZoom = .MinZoom
            strMaxZoom = .MaxZoom
            strLayerName = .Name
            iAnswer = MsgBox("Layername: " & strLayerName & vbCrLf & "Min. zoom:  " & strMinZoom & vbCrLf & "Max. zoom:  " & strMaxZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objLayer
End Sub
