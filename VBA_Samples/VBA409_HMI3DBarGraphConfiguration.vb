Sub HMI3DBarGraphConfiguration()
'VBA409
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BaseY = 100
    End With
End Sub
