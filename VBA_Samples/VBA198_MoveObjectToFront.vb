Sub MoveObjectToFront()
'VBA198
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = True
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = False
    End With
    MsgBox "The objects circle and rectangle are created" & vbCrLf & "Only the circle is selected!"
    ActiveDocument.Selection.BringToFront
    MsgBox "The selection is moved to the front."
End Sub
