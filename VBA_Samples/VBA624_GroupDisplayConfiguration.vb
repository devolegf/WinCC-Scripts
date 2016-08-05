Sub GroupDisplayConfiguration()
'VBA624
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCGUTextColorOff = RGB(0, 0, 255)
    End With
End Sub
