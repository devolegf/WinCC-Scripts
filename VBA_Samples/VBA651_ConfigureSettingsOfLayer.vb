Sub ConfigureSettingsOfLayer()
'VBA651
    Dim objLayer As HMILayer
    Set objLayer = ActiveDocument.Layers(1)
    With objLayer
        'Configure "Layer 0"
        .MinZoom = 10
        .MaxZoom = 100
        .Name = "Configured with VBA"
    End With
    'Define fade-in and fade-out of objects:
    With ActiveDocument
        .LayerDecluttering = True
        .ObjectSizeDecluttering = True
        .SetDeclutterObjectSize 50, 100
    End With
End Sub
