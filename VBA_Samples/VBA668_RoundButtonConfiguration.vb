Sub RoundButtonConfiguration()
'VBA668
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDownReferenced = False
    End With
End Sub
