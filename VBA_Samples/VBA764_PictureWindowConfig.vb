Sub PictureWindowConfig()
'VBA764
     Dim objPicWindow As HMIPictureWindow
     Set objPicWindow = ActiveDocument.HMIObjects.AddHMIObject("PicWindow1", "HMIPictureWindow")
     With objPicWindow
          .UpdateCycle = 5
     End With
End Sub
