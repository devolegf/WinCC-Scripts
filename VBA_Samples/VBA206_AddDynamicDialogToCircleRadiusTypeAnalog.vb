Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA206
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
    End With
End Sub
