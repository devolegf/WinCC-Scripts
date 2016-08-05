Sub InsertDocumentSpecificToolbar()
'VBA352
    Dim objToolbar As HMIToolbar
    Set objToolbar = ActiveDocument.CustomToolbars.Add("d_Toolbar1")
End Sub
