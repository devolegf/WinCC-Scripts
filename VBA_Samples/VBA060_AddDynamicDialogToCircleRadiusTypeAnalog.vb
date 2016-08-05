Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA60
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
'
    'Create dynamic
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
'
    'Configure dynamic. "ResultType" defines the type of valuerange:
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub