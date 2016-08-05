Sub RoundButtonConfiguration()
'VBA670
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDownTransparent = RGB(255, 255, 0)
        .PicDownUseTransColor = True
    End With
End Sub
