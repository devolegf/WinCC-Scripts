Sub AddObject()
'VBA319
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMIObject", "HMICircle")
'
    'Standard properties (e.g. "Position") are available every time:
    objObject.Top = 40
    objObject.Left = 40
'
    'Individual properties have to be called using
    'property "Properties":
    objObject.Properties("FlashBackColor") = True
End Sub
