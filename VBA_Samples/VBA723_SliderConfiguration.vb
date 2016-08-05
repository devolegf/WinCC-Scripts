Sub SliderConfiguration()
'VBA723
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .SmallChange = 4
    End With
End Sub
