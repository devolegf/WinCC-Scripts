Sub AddCircle()
'VBA31
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMICircle", "HMICircle")
'
    'The same as in example 1, but here you can set/get direct the
    specific properties of the circle:
    objCircle.Top = 80
    objCircle.Left = 80
    objCircle.FlashBackColor = True
End Sub