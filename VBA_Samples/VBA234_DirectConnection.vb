Sub DirectConnection()
'VBA234
    Dim objButton As HMIButton
    Dim objRectangleA As HMIRectangle
    Dim objRectangleB As HMIRectangle
    Dim objEvent As HMIEvent
    Dim objDirConnection As HMIDirectConnection
    Set objRectangleA = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_A", "HMIRectangle")
    Set objRectangleB = ActiveDocument.HMIObjects.AddHMIObject("Rectangle_B", "HMIRectangle")
    Set objButton = ActiveDocument.HMIObjects.AddHMIObject("myButton", "HMIButton")
    With objRectangleA
        .Top = 100
        .Left = 100
    End With
    With objRectangleB
        .Top = 250
        .Left = 400
        .BackColor = RGB(255, 0, 0)
    End With
    With objButton
        .Top = 10
        .Left = 10
        .Width = 90
        .Height = 50
        .Text = "SetPosition"
    End With
'
    'Directconnection is initiated on mouseclick:
    Set objDirConnection = objButton.Events(1).Actions.AddAction(hmiActionCreationTypeDirectConnection)
    With objDirConnection
        'Sourceobject: Property "Top" of "Rectangle_A"
        .SourceLink.Type = hmiSourceTypeProperty
        .SourceLink.ObjectName = "Rectangle_A"
        .SourceLink.AutomationName = "Top"
'
        'Targetobject: Property "Left" of "Rectangle_B"
        .DestinationLink.Type = hmiDestTypeProperty
        .DestinationLink.ObjectName = "Rectangle_B"
        .DestinationLink.AutomationName = "Left"
    End With
End Sub
