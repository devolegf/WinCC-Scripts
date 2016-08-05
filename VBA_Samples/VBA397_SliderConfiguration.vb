Sub SliderConfiguration()
'VBA397
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .BackColorTop = RGB(255, 255, 0)
    End With
End Sub
