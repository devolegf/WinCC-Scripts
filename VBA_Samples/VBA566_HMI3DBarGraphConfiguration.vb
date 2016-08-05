Sub HMI3DBarGraphConfiguration()
'VBA566
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer04Checked = True
    End With
End Sub
