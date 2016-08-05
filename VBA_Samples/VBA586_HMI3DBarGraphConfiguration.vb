Sub HMI3DBarGraphConfiguration()
'VBA586
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
        .Layer10Value = 100
    End With
End Sub
