Sub CreateDocumentToolbars()
'VBA26
    'Declare toolbarobjects:
    Dim objToolbar1 As HMIToolbar
    Dim objToolbar2 As HMIToolbar
'
    'Insert toolbars ("Add"-Methode) with
    'Parameter - "Key":
    Set objToolbar1 = ActiveDocument.CustomToolbars.Add("DocToolbar1")
    Set objToolbar2 = ActiveDocument.CustomToolbars.Add("DocToolbar2")
End Sub