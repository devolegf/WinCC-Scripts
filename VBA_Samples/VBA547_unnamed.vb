Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
'VBA547
    Dim strClicked As String
    Dim objMenuItem As HMIMenuItem
    Set objMenuItem = MenuItem
'
    '"strClicked can get two values:
    '(1) "DeleteAllRectangles" and
    '(2) "DeleteAllCircles"
    strClicked = objMenuItem.Key
'
    'To analyse "strClicked" with "Select Case"
    Select Case strClicked
        Case "DeleteAllRectangles"
'
            'Instead of "MsgBox" a procedurecall (e.g. "Call <Prozedurname>") can stay here
            MsgBox "'Delete rectangle' was clicked"
        Case "DeleteAllCircles"
            MsgBox "'Delete Circles' was clicked"
    End Select
End Sub
