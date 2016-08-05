Sub SliderConfiguration()
'VBA396
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .BackColorBottom = RGB(0, 0, 255)
    End With
End Sub
