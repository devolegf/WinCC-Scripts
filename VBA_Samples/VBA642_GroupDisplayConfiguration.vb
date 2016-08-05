Sub GroupDisplayConfiguration()
'VBA642
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MessageClass = 0
        .MCGUBackColorOff = RGB(255, 0, 0)
    End With
End Sub
