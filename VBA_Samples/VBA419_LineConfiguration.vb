Sub LineConfiguration()
'VBA419
    Dim objLine As HMILine
    Set objLine = ActiveDocument.HMIObjects.AddHMIObject("Line1", "HMILine")
    With objLine
        .BorderEndStyle = 393219
    End With
End Sub
