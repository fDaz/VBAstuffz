Attribute VB_Name = "DepthExit"
Private ExPosR, ExPosC As Integer

Sub SetPosR(r)

ExPosR = r

End Sub

Sub SetPosC(c)

ExPosC = c

End Sub

Function GetPosR()

GetPosR = ExPosR

End Function

Function GetPosC()

GetPosC = ExPosC

End Function

Sub Draw()

ICSRH.Cells(ExPosR, ExPosC).Font.Color = vbBlack
ICSRH.Cells(ExPosR, ExPosC) = ">"

If ExPosR = PlayerChar.GetPosR And ExPosC = PlayerChar.GetPosC Then Call MessageLog.NewMessage("You see a set of stairs heading downwards.")

End Sub

