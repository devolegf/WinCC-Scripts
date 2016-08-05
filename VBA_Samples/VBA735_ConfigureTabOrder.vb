Sub ConfigureTabOrder()
'VBA735
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub
