Sub SliderConfiguration()
'VBA459
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ColorTop = RGB(255, 128, 0)
    End With
End Sub
