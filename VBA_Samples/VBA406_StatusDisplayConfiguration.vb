Sub StatusDisplayConfiguration()
'VBA406
    Dim objStatDisp As HMIStatusDisplay
    Set objStatDisp = ActiveDocument.HMIObjects.AddHMIObject("Statusdisplay1", "HMIStatusDisplay")
    With objStatDisp
'
        'To use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "Testpicture.BMP" with the name of
        'the picture you copied
        .BasePicture = "Testpicture.BMP"
    End With
End Sub
