Sub AddDynamicDialogToCircleRadiusTypeBit()
'VBA216
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_B", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeBit
        .BitResultInfo.BitNumber = 1
        .BitResultInfo.BitSetValue = 40
        .BitResultInfo.BitNotSetValue = 80
    End With
End Sub
