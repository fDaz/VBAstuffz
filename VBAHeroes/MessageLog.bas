Attribute VB_Name = "MessageLog"
Sub Init()

ICSRH.Range(Cells(7, 57), Cells(26, 60)).HorizontalAlignment = xlLeft

End Sub

Sub NewMessage(msgstring)

For i = 7 To 25

    ICSRH.Cells(i, 57) = ICSRH.Cells(i + 1, 57)

Next

ICSRH.Cells(26, 57) = msgstring

End Sub

Sub AmendMessage(msgstring)

ICSRH.Cells(26, 57) = ICSRH.Cells(26, 57) & msgstring

End Sub
