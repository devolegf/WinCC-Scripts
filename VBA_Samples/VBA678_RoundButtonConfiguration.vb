Sub RoundButtonConfiguration()
'VBA678
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicUpTransparent = RGB(0, 0, 255)
        .PicUpUseTransColor = True
    End With
End Sub
