Sub HMI3DBarGraphConfiguration()
'VBA584
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
    End With
End Sub
