Sub AddDynamicAsVariableDirectToProperty()
'VBA141
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("MyCircle", "HMICircle")
    'Make property "Top" dynamic:
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic")
'
    'Define cycle-time
    With objVariableTrigger
        .CycleType = hmiCycleType_2s
    End With
End Sub
