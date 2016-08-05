Sub LDLabelInfo()
'VBA592
    Dim colLayerLngTexts As HMILanguageTexts
    Dim objLayerLngText As HMILanguageText
    Dim iIndex As Integer
    Dim iAnswer As Integer
    Dim strResult As String
    iIndex = 1
    For iIndex = 1 To ActiveDocument.Layers.Count
'
        'Save all labels of layers into collection of "colLayerLngTexts":
        Set colLayerLngTexts = ActiveDocument.Layers(iIndex).LDNames
        For Each objLayerLngText In colLayerLngTexts
            strResult = strResult & vbCrLf & objLayerLngText.LanguageID & " - " & objLayerLngText.DisplayName
        Next objLayerLngText
        iAnswer = MsgBox(strResult, vbOKCancel)
        strResult = ""
        If vbCancel = iAnswer Then Exit For
    Next iIndex
End Sub
