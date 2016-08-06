VBS111
Dim lngFactor
Dim dblAxisX
Dim dblAxisY
Dim objTrendControl
Set objTrendControl = ScreenItems("Control1")
For lngFactor = -100 To 100
    dblAxisX = CDbl(lngFactor * 0.02)
    dblAxisY = CDbl(dblAxisX * dblAxisX + 2 * dblAxisX + 1)
    objTrendControl.DataX = dblAxisX
    objTrendControl.DataY = dblAxisY
    objTrendControl.InsertData = True
Next
