Sub GroupDisplayConfiguration()
'VBA627
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOBackColorOff = RGB(255, 0, 0)
    End With
End Sub
