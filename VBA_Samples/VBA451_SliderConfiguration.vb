Sub SliderConfiguration()
'VBA451
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ColorBottom = RGB(255, 0, 0)
    End With
End Sub
