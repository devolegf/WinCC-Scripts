Sub AddDynamicDialogToCircleRadiusTypeAnalog()
'VBA775
    Dim objDynDialog As HMIDynamicDialog
    Dim objCircle As HMICircle
    Set objCircle = ActiveDocument.HMIObjects.AddHMIObject("Circle_A", "HMICircle")
    Set objDynDialog = objCircle.Radius.CreateDynamic(hmiDynamicCreationTypeDynamicDialog, "'NewDynamic1'")
    With objDynDialog
        .ResultType = hmiResultTypeAnalog
        .AnalogResultInfos.ElseCase = 200
'
        'Activate analysis of variablestate
        .VariableStateChecked = True
    End With
    With objDynDialog.VariableStateValues(1)
'
        'define a value for every state:
        .VALUE_ACCESS_FAULT = 20
        .VALUE_ADDRESS_ERROR = 30
        .VALUE_CONVERSION_ERROR = 40
        .VALUE_HANDSHAKE_ERROR = 60
        .VALUE_HARDWARE_ERROR = 70
        .VALUE_INVALID_KEY = 80
        .VALUE_MAX_LIMIT = 90
        .VALUE_MAX_RANGE = 100
        .VALUE_MIN_LIMIT = 110
        .VALUE_MIN_RANGE = 120
        .VALUE_NOT_ESTABLISHED = 130
        .VALUE_SERVERDOWN = 140
        .VALUE_STARTUP_VALUE = 150
        .VALUE_TIMEOUT = 160
    End With
End Sub
