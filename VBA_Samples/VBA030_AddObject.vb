Sub AddObject()
'VBA30
    Dim objObject As HMIObject
    Set objObject = ActiveDocument.HMIObjects.AddHMIObject("CircleAsHMIObject", "HMICircle")
'
    'standard-properties (e.g. the position) are available every time:
    objObject.Top = 40
    objObject.Left = 40
'
    'non-standard properties can be accessed using the Properties-collection:
    objObject.Properties("FlashBackColor") = True
End Sub
