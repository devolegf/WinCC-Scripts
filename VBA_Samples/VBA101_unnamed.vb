Private Sub Document_MenuItemClicked(ByVal MenuItem As IHMIMenuItem)
'VBA101
    Dim objMenuItem As HMIMenuItem
    Dim varMenuItemKey As Variant
    Set objMenuItem = MenuItem
'
    '"objMenuItem" contains the clicked menu-item
    '"varMenuItemKey" contains the value of parameter "Key"
    'from the clicked userdefined menu-item
    varMenuItemKey = objMenuItem.Key
    Select Case MenuItemKey
        Case "mItem1_1"
            MsgBox "The first menu-item was clicked!"
    End Select
End Sub