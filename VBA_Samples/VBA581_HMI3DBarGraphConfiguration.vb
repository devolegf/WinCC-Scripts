Sub HMI3DBarGraphConfiguration()
'VBA581
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer09Checked = True
    End With
End Sub
