Sub EditStatusDisplay()
'VBA341
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects("Statusdisplay1")
    objStatusDisplay.BorderColor = RGB(255, 0, 0)
End Sub
