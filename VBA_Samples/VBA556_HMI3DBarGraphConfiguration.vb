Sub HMI3DBarGraphConfiguration()
'VBA556
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer00Checked = True
        .Layer00Value = 0
    End With
End Sub
