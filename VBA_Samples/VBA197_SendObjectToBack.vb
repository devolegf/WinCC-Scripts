Sub SendObjectToBack()
'VBA197
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
    MsgBox "The objects circle and rectangle are created" & vbCrLf & "Only the rectangle is selected!"
    ActiveDocument.Selection.SendToBack
    MsgBox "The selection is moved to the back."
End Sub
