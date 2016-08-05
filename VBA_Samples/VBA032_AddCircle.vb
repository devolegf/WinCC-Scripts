Sub AddCircle()
'VBA32
    'Creates object of type "HMICircle"
    Dim objCircle As HMICircle
'
    'Add object in active document
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("My Circle", "HMICircle")
End Sub