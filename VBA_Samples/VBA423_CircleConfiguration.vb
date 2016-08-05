Sub CircleConfiguration()
'VBA423
    Dim objCircle As IHMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    With objCircle
        .BorderWidth = 2
    End With
End Sub
