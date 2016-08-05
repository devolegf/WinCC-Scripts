Sub DynamicToRadiusOfNewCircle()
'VBA318
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects("Circle")
    Set objVariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
    objVariableTrigger.CycleType = hmiCycleType_2s
End Sub
