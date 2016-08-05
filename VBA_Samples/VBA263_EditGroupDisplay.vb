Sub EditGroupDisplay()
'VBA263
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects("Groupdisplay")
    objGroupDisplay.BackColor = RGB(255, 0, 0)
End Sub
