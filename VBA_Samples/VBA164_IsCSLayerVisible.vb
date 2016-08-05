Sub IsCSLayerVisible()
'VBA164
    Dim objView As HMIView
    Dim strLayerName As String
    Dim iLayerIdx As Integer
    Set objView = ActiveDocument.Views(1)
    objView.Activate
    iLayerIdx = 2
    strLayerName = ActiveDocument.Layers(iLayerIdx).Name
    If objView.IsCSLayerVisible(iLayerIdx) = True Then
        MsgBox "CS " & strLayerName & " is visible"
    Else
        MsgBox "CS " & strLayerName & " is invisible"
    End If
End Sub
