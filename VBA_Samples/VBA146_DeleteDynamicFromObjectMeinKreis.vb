Sub DeleteDynamicFromObjectMeinKreis()
'VBA146
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("MyCircle")
    objCircle.Top.DeleteDynamic
End Sub
