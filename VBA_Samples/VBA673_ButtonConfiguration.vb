Sub ButtonConfiguration()
'VBA673
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
'
        'Toi use this example copy a Bitmap-Graphic
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturename "TestPicture1.BMP" with the name of
        'the picture you copied
        .PictureDeactivated = "TestPicture1.BMP"
    End With
End Sub
