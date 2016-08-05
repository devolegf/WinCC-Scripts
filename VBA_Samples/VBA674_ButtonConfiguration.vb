Sub ButtonConfiguration()
'VBA674
    Dim objButton As HMIButton
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("Button1", "HMIButton")
    With objButton
    '
        'To use this example copy two Bitmap-Graphics
        'to the "GraCS"-Folder of the actual project.
        'Replace the picturenames "TestPicture1.BMP" and "TestPicture2.BMP"
        'with the names of the pictures you copied
        .PictureDown = "TestPicture1.BMP"
        .PictureUp = "TestPicture2.BMP"
    End With
End Sub
