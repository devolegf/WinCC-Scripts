Sub AddDynamicDialogToCircleRadiusTypeBinary()
'VBA699
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_C", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBool
        .BinaryResultInfo.NegativeValue = 20
        .BinaryResultInfo.PositiveValue = 40
    End With
End Sub
