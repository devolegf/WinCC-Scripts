Sub StatusDisplayConfiguration()
'VBA505
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashPicTransColor = RGB(255, 255, 0)
        .FlashPicUseTransColor = True
    End With
End Sub
