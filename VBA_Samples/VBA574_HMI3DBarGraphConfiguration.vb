Sub HMI3DBarGraphConfiguration()
'VBA574
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer06Checked = True
        .Layer06Value = 60
    End With
End Sub
