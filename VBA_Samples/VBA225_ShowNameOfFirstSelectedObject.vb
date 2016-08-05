Sub ShowNameOfFirstSelectedObject()
'VBA225
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name from the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub
