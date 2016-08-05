Sub DeleteActionOfCircleAB()
'VBA144
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle_AB")
    objCircle.Radius.Events(1).Actions(1).Delete
End Sub
