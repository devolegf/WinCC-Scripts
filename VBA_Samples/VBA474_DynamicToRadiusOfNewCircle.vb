Sub DynamicToRadiusOfNewCircle()
'VBA474
    Dim objCircle As hmiCircle
    Dim VariableTrigger As HMIVariableTrigger
    Set objCircle = Application.ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set VariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
    VariableTrigger.CycleType = hmiVariableCycleType_2s
End Sub
