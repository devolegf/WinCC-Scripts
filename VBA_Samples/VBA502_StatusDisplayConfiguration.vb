Sub StatusDisplayConfiguration()
'VBA502
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
        .FlashPicReferenced = True
    End With
End Sub
