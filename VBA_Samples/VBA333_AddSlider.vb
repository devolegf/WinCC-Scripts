Sub AddSlider()
'VBA333
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects.AddHMIObject("Slider1", "HMISlider")
End Sub
