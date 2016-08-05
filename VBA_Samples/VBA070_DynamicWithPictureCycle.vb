Sub DynamicWithPictureCycle()
'VBA70
    Dim objVBScript As HMIScriptInfo
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_Picture", "HMICircle")
    Set objVBScript = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeVBScript)
    With objVBScript
        .Trigger.Type = hmiTriggerTypePictureCycle
        .Trigger.Name = "VBA_PictureCycle"
        .SourceCode = ""
    End With
End Sub