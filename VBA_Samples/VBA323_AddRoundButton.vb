Sub AddRoundButton()
'VBA323
    Dim objRoundButton As HMIRoundButton
    Set objRoundButton = ActiveDocument.HMIObjects.AddHMIObject("Roundbutton1", "HMIRoundButton")
End Sub
