Sub GroupDisplayConfiguration()
'VBA630
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKOTextColorOff = RGB(0, 0, 255)
    End With
End Sub
