Sub HMI3DBarGraphConfiguration()
'VBA559
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer01Checked = True
        .Layer01Value = 10
    End With
End Sub
