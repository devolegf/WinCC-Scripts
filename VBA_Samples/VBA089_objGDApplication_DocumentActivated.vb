Private Sub objGDApplication_DocumentActivated(ByVal Document As IHMIDocument)
'VBA89
    MsgBox "The document " & Document.Name & " got the focus." & vbCrLf & "This event is raised by the application."
End Sub