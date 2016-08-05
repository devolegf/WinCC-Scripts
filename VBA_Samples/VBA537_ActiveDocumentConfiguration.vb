Sub ActiveDocumentConfiguration()
'VBA537
    Application.ActiveDocument.Views.Add
    'If you comment out the following line
    'and recall the procedure, the output of
    'the messagebox is different
    Application.ActiveDocument.Views(1).Activate
'
    'Output state of copy:
    MsgBox Application.ActiveDocument.Views(1).IsActive
End Sub
