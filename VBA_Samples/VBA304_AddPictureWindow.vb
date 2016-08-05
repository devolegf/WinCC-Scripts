Sub AddPictureWindow()
'VBA304
    Dim objPictureWindow As HMIPictureWindow
    Set objPictureWindow = ActiveDocument.HMIObjects.AddHMIObject("PictureWindow1", "HMIPictureWindow")
End Sub
