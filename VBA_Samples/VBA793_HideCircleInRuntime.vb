Sub HideCircleInRuntime()
'VBA793
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("myCircle", "HMICircle")
    objCircle.Visible = False
End Sub