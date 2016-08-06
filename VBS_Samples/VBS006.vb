Sub OnClick(ByVal Item)
'VBS6
    Dim lngAnswer
    Dim lngIndex
    lngIndex = 1
    For lngIndex = 1 To ScreenItems.Count
        lngAnswer = MsgBox(ScreenItems(lngIndex).Objectname, vbOKCancel)
        If vbCancel = lngAnswer Then Exit For
    Next
End Sub