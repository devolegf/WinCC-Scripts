Sub CreateApplicationToolbars()
'VBA6
    'Declare toolbar-objects...:
    Dim objToolbar1 As HMIToolbar
    Dim objToolbar2 As HMIToolbar
'
    'Add the toolbars with parameter "Key"
    Set objToolbar1 = Application.CustomToolbars.Add("AppToolbar1")
    Set objToolbar2 = Application.CustomToolbars.Add("AppToolbar2")
    
End Sub