Sub DeleteToolbar()
'VBA349
    Dim objToolbar As HMIToolbar
    Set objToolbar = ActiveDocument.CustomToolbars(1)
    objToolbar.Delete
End Sub
