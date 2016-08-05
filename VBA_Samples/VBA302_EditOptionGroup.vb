Sub EditOptionGroup()
'VBA302
    Dim objOptionGroup As HMIOptionGroup
    Set objOptionGroup = ActiveDocument.HMIObjects("Radio-Box")
    objOptionGroup.BorderColor = RGB(255, 0, 0)
End Sub
