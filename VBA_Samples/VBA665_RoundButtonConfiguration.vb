Sub RoundButtonConfiguration()
'VBA665
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("RButton1", "HMIRoundButton")
    With objRoundButton
        .PicDeactReferenced = False
    End With
End Sub
