Sub HMI3DBarGraphConfiguration()
'VBA585
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer10Checked = True
        .Layer10Color = RGB(255, 0, 255)
    End With
End Sub
