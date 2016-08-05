Sub StatusDisplayConfiguration()
'VBA500
    Dim objsDisplay As HMIStatusDisplay
    Set objsDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objsDisplay
        .FlashFlashPicture = True
    End With
End Sub
