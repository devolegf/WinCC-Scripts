Sub ShowFirstObjectOfCollection()
'VBA317
    Dim objCircle As HMICircle
    Dim strName As String
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle", "HMICircle")
    strName = objCircle.Properties(1).Name
    MsgBox strName
End Sub
