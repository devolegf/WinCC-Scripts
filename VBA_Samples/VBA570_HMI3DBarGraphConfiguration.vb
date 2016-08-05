Sub HMI3DBarGraphConfiguration()
'VBA570
    Dim obj3DBar As HMI3DBarGraph
    Set obj3DBar = ActiveDocument.HMIObjects.AddHMIObject("3DBar1", "HMI3DBarGraph")
    With obj3DBar
        .Layer05Checked = True
        .Layer05Color = RGB(255, 0, 255)
    End With
End Sub
