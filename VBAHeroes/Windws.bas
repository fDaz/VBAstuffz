Attribute VB_Name = "Windws"
Sub InitWindow(startR, startC, endR, endC)

With ICSRH.Range(Cells(startR, startC), Cells(endR, endC))
    .Interior.Color = vbBlack
    .Font.Color = vbWhite
    .Font.Size = 12
End With

ICSRH.Range(Cells(startR + 1, startC + 1), Cells(endR - 1, endC - 1)).HorizontalAlignment = xlLeft

ICSRH.Range(Cells(startR, startC), Cells(endR, endC)) = "#"
ICSRH.Range(Cells(startR + 1, startC + 1), Cells(endR - 1, endC - 1)) = ""

End Sub

Sub CloseWindow(startR, startC, endR, endC)

With ICSRH.Range(Cells(startR, startC), Cells(endR, endC))
    .Font.Color = vbBlack
    .HorizontalAlignment = xlCenter
    .Font.Size = 14
End With

ICSRH.Range(Cells(startR, startC), Cells(endR, endC)) = ""

Call DepthMap.ReDrawCurrMap(startR, startC, endR, endC)
Call DepthMap.Refresh

Call ICSRH.SetControlType(0)

End Sub
