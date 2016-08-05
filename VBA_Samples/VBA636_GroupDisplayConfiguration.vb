Sub GroupDisplayConfiguration()
'VBA636
    Dim objGroupDisplay As HMIGroupDisplay
    Set objGroupDisplay = ActiveDocument.HMIObjects.AddHMIObject("GroupDisplay1", "HMIGroupDisplay")
    With objGroupDisplay
        .MCKQTextColorOff = RGB(0, 0, 255)
    End With
End Sub
