Sub EditCheckBox()
'VBA221
    Dim objCheckBox As HMICheckBox
    Set objCheckBox = ActiveDocument.HMIObjects("CheckBox")
    objCheckBox.BorderColor = RGB(255, 0, 0)
End Sub
