Sub HMI3DBarGraphConfiguration()
'VBA408
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .BaseX = 80
    End With
End Sub
