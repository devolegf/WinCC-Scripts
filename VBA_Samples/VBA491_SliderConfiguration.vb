Sub SliderConfiguration()
'VBA491
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("SliderObject1", "HMISlider")
    With objSlider
        .ExtendedOperation = True
    End With
End Sub
