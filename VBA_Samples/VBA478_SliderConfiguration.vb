Sub SliderConfiguration()
'VBA478
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .Direction = True
    End With
End Sub
