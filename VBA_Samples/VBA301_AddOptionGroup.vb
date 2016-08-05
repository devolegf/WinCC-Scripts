Sub AddOptionGroup()
'VBA301
    Dim objOptionGroup As HMIOptionGroup
    Set objOptionGroup = ActiveDocument.HMIObjects.AddHMIObject("Radio-Box", "HMIOptionGroup")
End Sub
