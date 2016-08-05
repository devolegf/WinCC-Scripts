Sub StatusDisplayConfiguration()
'VBA509
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashRateFlashPic = 1
    End With
End Sub
