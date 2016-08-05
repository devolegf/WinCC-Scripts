Sub HMI3DBarGraphConfiguration()
'VBA560
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer02Checked = True
    End With
End Sub
