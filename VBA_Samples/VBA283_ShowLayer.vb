Sub ShowLayer()
'VBA283
    Dim colLayers As HMILayers
    Dim objLayer As HMILayer
    Dim strLayerList As String
    Dim iCounter As Integer
    iCounter = 1
    Set colLayers = ActiveDocument.Layers
    For Each objLayer In colLayers
        If 1 = iCounter Mod 2 And 32 > iCounter Then
            strLayerList = strLayerList & vbCrLf
        ElseIf 11 > iCounter Then
            strLayerList = strLayerList & "       "
        Else
            strLayerList = strLayerList & "     "
        End If
        strLayerList = strLayerList & objLayer.Name
        iCounter = iCounter + 1
    Next objLayer
    MsgBox strLayerList
End Sub
