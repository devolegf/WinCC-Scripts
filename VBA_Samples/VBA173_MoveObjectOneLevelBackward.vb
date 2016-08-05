Sub MoveObjectOneLevelBackward()
'VBA173
    Dim objCircle As HMICircle
    Dim objRectangle As HMIRectangle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("sCircle", "HMICircle")
    Set objRectangle = ActiveDocument.HMIObjects.AddHMIObject("sRectangle", "HMIRectangle")
    With objCircle
        .Top = 40
        .Left = 40
        .Selected = False
    End With
    With objRectangle
        .Top = 40
        .Left = 40
        .Width = 100
        .Height = 50
        .BackColor = RGB(255, 0, 255)
        .Selected = True
    End With
    MsgBox "Objects created and selected!"
    ActiveDocument.Selection.BackwardOneLevel
End Sub
