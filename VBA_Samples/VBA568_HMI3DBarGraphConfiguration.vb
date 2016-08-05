Sub HMI3DBarGraphConfiguration()
'VBA568
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
        .Layer04Value = 40
    End With
End Sub
