Sub ConfigureTabOrder()
'VBA736
    With ActiveDocument
        .TABOrderAllHMIObjects = True
        .TABOrderKeyboard = False
        .TABOrderMouse = False
        .TABOrderOtherAction = False
    End With
End Sub
