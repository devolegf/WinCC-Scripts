Sub HMI3DBarGraphConfiguration()
'VBA583
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer09Checked = True
        .Layer09Value = 90
    End With
End Sub
