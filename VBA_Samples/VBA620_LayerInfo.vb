Sub LayerInfo()
'VBA620
    Dim colLayers As HMILayers
    Dim objSingleLayer As HMILayer
    Dim iAnswer As Integer
    Set colLayers = ActiveDocument.Layers
    For Each objSingleLayer In colLayers
        With objSingleLayer
            iAnswer = MsgBox("Layername: " & .Name & vbCrLf & "Min. zoom:  " & .MinZoom & vbCrLf & "Max. zoom:  " & .MaxZoom, vbOKCancel)
        End With
        If vbCancel = iAnswer Then Exit For
    Next objSingleLayer
End Sub
