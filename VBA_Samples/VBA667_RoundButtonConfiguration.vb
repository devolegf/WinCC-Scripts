Sub RoundButtonConfiguration()
'VBA667
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDeactTransparent = RGB(255, 0, 0)
        .PicDeactUseTransColor = True
    End With
End Sub
