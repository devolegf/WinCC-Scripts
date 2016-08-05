Sub ShowCustomToolbarsOfDocument()
'VBA350
    Dim colToolbars As HMIToolbars
    Dim objToolbar As HMIToolbar
    Dim strToolbarList As String
    Set colToolbars = ActiveDocument.CustomToolbars
    If 0 <> colToolbars.Count Then
        For Each objToolbar In colToolbars
            strToolbarList = strToolbarList & objToolbar.Key & vbCrLf
        Next objToolbar
    Else
        strToolbarList = "No toolbars existing"
    End If
    MsgBox strToolbarList
End Sub
