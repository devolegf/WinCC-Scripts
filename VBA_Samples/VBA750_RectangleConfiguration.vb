Sub RectangleConfiguration()
'VBA750
    Dim objRectangle As HMIRectangle
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("Rectangle1", "HMIRectangle")
    With objRectangle
        MsgBox "Objecttype: " & .Type
    End With
End Sub
