Sub AddStaticText()
'VBA337
    Dim objStaticText As HMIStaticText
    Set objStaticText = ActiveDocument.HMIObjects.AddHMIObject("Static_Text1", "HMIStaticText")
End Sub
