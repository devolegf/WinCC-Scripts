Sub EditHMIObject()
'VBA34
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("myCircleAsObject", "HMICircle")
    With objObject
        'Using object's properties collection to access its specific properties
        .Properties("BorderWidth") = 4
        .Properties("BorderColor") = RGB(255, 0, 0)
    End With
End Sub
