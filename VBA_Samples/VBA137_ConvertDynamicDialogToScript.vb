Sub ConvertDynamicDialogToScript()
'VBA137
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
'
    'Create dynamic
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
'
    'configure dynamic. "ResultType" defines the valuerange-type:
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.Add 50, 40
        .AnalogResultInfos.Add 100, 80
        .AnalogResultInfos.ElseCase = 100
        MsgBox "The dynamic-dialog will be changed into a C-script."
        .ConvertToScript
    End With
End Sub
