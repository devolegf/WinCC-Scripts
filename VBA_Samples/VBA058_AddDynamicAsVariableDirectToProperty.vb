Sub AddDynamicAsVariableDirectToProperty()
'VBA58
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
    'Create dynamic at property "Top"
    Set objVariableTrigger = objCircle.Top.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'define cycle-time
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub
