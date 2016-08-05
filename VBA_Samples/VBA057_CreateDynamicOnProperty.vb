Sub CreateDynamicOnProperty()
'VBA57
    Dim objVariableTrigger As HMIVariableTrigger
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle1", "HMICircle")
'
    'Create dynamic with type "direct Variableconnection" at the
    'property "Radius":
    Set objVariableTrigger = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVariableDirect, "NewDynamic1")
'
    'To complete dynamic, e.g. define cycle:
    With objVariableTrigger
        .CycleType = hmiVariableCycleType_2s
    End With
End Sub
