Sub AddCheckBox()
'VBA220
    Dim objCheckBox As HMICheckBox
    Set objCheckBox = ActiveDocument.HMIObjects.AddHMIObject("CheckBox", "HMICheckBox")
End Sub
