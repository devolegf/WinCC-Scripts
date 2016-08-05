Sub AddGroupDisplay()
'VBA262
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("Groupdisplay", "HMIGroupDisplay")
End Sub
