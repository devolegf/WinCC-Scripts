Sub PictureWindowConfig()
'VBA710
    Dim objPicWindow As HMIPictureWindow
    Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
    With objPicWindow
        .AdaptPicture = False
        .AdaptSize = False
        .Caption = True
        .CaptionText = "Picturewindow in runtime"
        .OffsetLeft = 5
        .OffsetTop = 10
        'Replace the picturename "Test.PDL" with the name of
        'an existing document from your "GraCS"-Folder of your active project
        .PictureName = "Test.PDL"
        .ScrollBars = True
        .ServerPrefix = ""
        .TagPrefix = "Struct."
        .UpdateCycle = 5
        .Zoom = 100
    End With
End Sub
