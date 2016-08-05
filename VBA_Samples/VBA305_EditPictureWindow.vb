Sub EditPictureWindow()
'VBA305
    Dim objPictureWindow As HMIPictureWindow
    Set objPictureWindow = ActiveDocument.HMIObjects("PictureWindow1")
    objPictureWindow.Sizeable = True
End Sub
