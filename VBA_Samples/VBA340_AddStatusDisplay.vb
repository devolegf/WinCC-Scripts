Sub AddStatusDisplay()
'VBA340
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
End Sub
