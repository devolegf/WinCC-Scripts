Sub ShowNameOfFirstSelectedObject()
'VBA325
    'Select all objects in the picture:
    ActiveDocument.Selection.SelectAll
    'Get the name of the first object of the selection:
    MsgBox ActiveDocument.Selection(1).ObjectName
End Sub
