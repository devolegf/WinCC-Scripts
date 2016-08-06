Private Sub Document_ToolbarItemClicked(ByVal ToolbarItem As IHMIToolbarItem)
'VBA108
    Dim objToolbarItem As HMIToolbarItem
    Dim varToolbarItemKey As Variant
    Set objToolbarItem = ToolbarItem
'
    '"varToolbarItemKey" contains the value of parameter "Key"
    'from the clicked userdefined toolbar-item
    varToolbarItemKey = objToolbarItem.Key
'
    Select Case varToolbarItemKey
        Case "tItem1_1"
            MsgBox "The first Toolbar-Icon was clicked!"
    End Select
End Sub
