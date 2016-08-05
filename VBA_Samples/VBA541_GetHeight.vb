Sub GetHeight()
'VBA541
    Dim objGroup As HMIGroup
    'Next line uses the property "Item" to get a group by name
    Set objGroup = ActiveDocument.HMIObjects.Item("Group1")
    'Otherwise next line uses index to identify a groupobject
    MsgBox "The height of object 2 is: " & objGroup.GroupedHMIObjects.Item(2).Height
End Sub
