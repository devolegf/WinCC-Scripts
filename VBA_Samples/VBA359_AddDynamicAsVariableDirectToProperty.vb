Sub AddDynamicAsVariableDirectToProperty()
'VBA359
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'Define cycletime
    With objVariableTrigger
        .CycleType = hmiCycleType_2s
    End With
End Sub
