Sub AddDynamicAsVariableIndirectToProperty()
'VBA59
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
 
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle2", "HMICircle")
    'Create dynamic on property "Left":
    Set objVariableTrigger = objCircle.Left.CreateDynamic(hmiDynamicCreationTypeVariableIndirect, "'NewDynamic1'")
'
    'Define cycle-time
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub
