Sub SelectObject()
'VBA35
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("mySelectedCircle", "HMICircle")
    ActiveDocument.HMIObjects("mySelectedCircle").Selected = True
End Sub
