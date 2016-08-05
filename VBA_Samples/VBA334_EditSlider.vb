Sub EditSlider()
'VBA334
    Dim objSlider As HMISlider
    Set objSlider = ActiveDocument.HMIObjects("Slider1")
    objSlider.ButtonColor = RGB(255, 0, 0)
End Sub
