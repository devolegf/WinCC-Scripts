Sub EditDefinedObjectType()
'VBA33
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircleAsCircle", "HMICircle")
    With objCircle
        'direct access of objectproperties available
        .BorderWidth = 4
        .BorderColor = RGB(255, 0, 255)
    End With
End Sub
