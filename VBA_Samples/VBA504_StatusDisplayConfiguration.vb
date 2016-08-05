Sub StatusDisplayConfiguration()
'VBA504
    Dim objStatusDisplay As HMIStatusDisplay
    Set objStatusDisplay = ActiveDocument.HMIObjects.AddHMIObject("StatusDisplay1", "HMIStatusDisplay")
    With objStatusDisplay
'
        'To use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "Testpicture.BMP" with the name of
        'the picture you copied
        .FlashPicture = "Testpicture.BMP"
    End With
End Sub
