Sub RTLayerVisibility()
'VBA165
    Dim strLayerName As String
    Dim iLayerIdx As Integer
    iLayerIdx = 2
    strLayerName = ActiveDocument.Layers(iLayerIdx).Name
    If ActiveDocument.IsRTLayerVisible(iLayerIdx) = True Then
        MsgBox "RT " & strLayerName & " is visible"
    Else
        MsgBox "RT " & strLayerName & " is invisible"
    End If
End Sub
