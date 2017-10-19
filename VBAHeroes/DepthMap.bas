Attribute VB_Name = "DepthMap"
Private MapOutLine(30, 55) '-1 unseen wall, 0: seen wall, 1: unseen floor, 2 seen space, 3 visible space
Private ObjPresent As New Collection
Private ActPresent As New Collection
Private Objkeys(10) As Integer

Sub Init()

Set ObjPresent = New Collection

End Sub
Sub Refresh()

'Redraws everything

Call CheckVisible

Call ReDrawCurrMap(PlayerChar.GetPosR - 6, PlayerChar.GetPosC - 6, PlayerChar.GetPosR + 6, PlayerChar.GetPosC + 6)

ICSRH.Range(Cells(2, 2), Cells(30, 55)) = ""

If MapOutLine(DepthExit.GetPosR, DepthExit.GetPosC) >= 2 Then
    DepthExit.Draw
End If

For i = 1 To ObjPresent.Count
    If MapOutLine(ObjPresent(i).PosR, ObjPresent(i).PosC) = 3 Then
        ICSRH.Cells(ObjPresent(i).PosR, ObjPresent(i).PosC).Font.Color = RGB(ObjPresent(i).Color("R"), ObjPresent(i).Color("G"), ObjPresent(i).Color("B"))
        ICSRH.Cells(ObjPresent(i).PosR, ObjPresent(i).PosC) = ObjPresent(i).Icon
    End If
Next

PlayerChar.Draw

End Sub

Sub StuffAtPlayerPos()

Dim firstmsg As Boolean
firstmsg = True

For i = 1 To ObjPresent.Count
    If ObjPresent(i).PosR = PlayerChar.GetPosR And ObjPresent(i).PosC = PlayerChar.GetPosC Then
        If firstmsg = True Then
            Call MessageLog.NewMessage("You see: ")
            firstmsg = False
        Else
            MessageLog.AmendMessage (", ")
        End If
        Call MessageLog.AmendMessage(ObjPresent(i).Name)
    End If
Next

End Sub

Sub PlaceItem(stuff As Object)

ObjPresent.Add stuff

End Sub
Sub PlaceWeapon(rowin, colin, ItemIDin, MatIDin)

Dim tItem As New cWeapon

tItem.PosR = rowin
tItem.PosC = colin
tItem.ItemID = ItemIDin
tItem.MatID = MatIDin

ObjPresent.Add tItem

End Sub

Sub PlaceArmor(rowin, colin, ItemIDin, MatIDin)

Dim tItem As New cArmor

tItem.PosR = rowin
tItem.PosC = colin
tItem.ItemID = ItemIDin
tItem.MatID = MatIDin

ObjPresent.Add tItem

End Sub

Sub PickUpItem()

Dim Count, winsize As Integer

For i = 0 To 10
    Objkeys(i) = 0
Next
Count = 0

For i = 1 To ObjPresent.Count
    If ObjPresent(i).PosR = PlayerChar.GetPosR And ObjPresent(i).PosC = PlayerChar.GetPosC Then
        Count = Count + 1
        Objkeys(Count) = i
    End If
Next

If Inventory.GetInvSize >= 15 Then
    Call MessageLog.NewMessage("Not enough room in the inventory")
    Exit Sub
End If

If Count = 1 Then
    Call Inventory.AddToInventory(ObjPresent(Objkeys(1)))
    Call MessageLog.NewMessage("Got: " & ObjPresent(Objkeys(1)).Name)
    ObjPresent.Remove (Objkeys(1))
    ICSRH.IncRounds
    Exit Sub
ElseIf Count > 1 Then
    winsize = (Count / 2) + 2
    Call Windws.InitWindow(16 - winsize - 1, 18, 16 + winsize, 39)
    ICSRH.Cells(16 - winsize, 20) = "Select an item to pick up:"
    For i = 1 To Count
        ICSRH.Cells(16 - winsize + 1 + i, 20) = GetKey(i) & ")"
        ICSRH.Cells(16 - winsize + 1 + i, 21) = ObjPresent(Objkeys(i)).Name
    Next
    ICSRH.Cells(16 + winsize - 1, 20) = "z)"
    ICSRH.Cells(16 + winsize - 1, 21) = "Exit"
    Call ICSRH.SetControlType(3)
    Exit Sub
End If

Call MessageLog.NewMessage("There is nothing here to pick up.")

End Sub

Sub SelectWhichItemToPickUp(numb)
'Assumption: "pick up" window is open, no changes in Objkeys() has been made

If Objkeys(numb) = 0 Then Exit Sub

Call Inventory.AddToInventory(ObjPresent(Objkeys(numb)))
Call MessageLog.NewMessage("Got: " & ObjPresent(Objkeys(numb)).Name)
ObjPresent.Remove (Objkeys(numb))

ICSRH.IncRounds
Call Windws.CloseWindow(4, 14, 28, 43)

End Sub

Function GetTile(r, c)

If r <= 30 And c <= 55 Then
    GetTile = MapOutLine(r, c)
Else
    MsgBox "Map coordinates " & r & " and " & c & " do not exist!"
End If

End Function

Sub HideMap()

Call SaveCurrMap
ICSRH.Range(Cells(2, 2), Cells(30, 55)).Interior.Color = vbBlack
ICSRH.Range(Cells(2, 2), Cells(30, 55)) = ""

End Sub

Sub SaveCurrMap()

For r = 1 To 30
    For c = 1 To 55
        If ICSRH.Cells(r, c) = "" Then MapOutLine(r, c) = -1 Else MapOutLine(r, c) = 1
    Next
Next

End Sub

Sub CheckVisible()

Dim cr, cc, resetRs, resetCs, resetRe, resetCe As Integer
Dim t As Double

resetRs = PlayerChar.GetPosR - 6
resetRe = PlayerChar.GetPosR + 6
resetCs = PlayerChar.GetPosC - 6
resetCe = PlayerChar.GetPosC + 6

If resetRs < 2 Then resetRs = 2
If resetCs < 2 Then resetCs = 2
If resetRe > 30 Then resetRe = 30
If resetCe > 55 Then resetCe = 55

'first reset the visible squares to seen squares
For i1 = resetRs To resetRe
    For i2 = resetCs To resetCe
        If MapOutLine(i1, i2) = 3 Then MapOutLine(i1, i2) = 2
    Next
Next

For t = 0 To 6.2 Step (90 / 7) * (3.14 / 180)
    cc = Round(PlayerChar.GetPosC + 5 * Cos(t))
    cr = Round(PlayerChar.GetPosR + 5 * Sin(t))
    Call CheckLine(PlayerChar.GetPosR, PlayerChar.GetPosC, cr, cc)
Next

Call CheckLine(PlayerChar.GetPosR, PlayerChar.GetPosC, PlayerChar.GetPosR + 3, PlayerChar.GetPosC + 3)
Call CheckLine(PlayerChar.GetPosR, PlayerChar.GetPosC, PlayerChar.GetPosR + 3, PlayerChar.GetPosC - 3)
Call CheckLine(PlayerChar.GetPosR, PlayerChar.GetPosC, PlayerChar.GetPosR - 3, PlayerChar.GetPosC - 3)
Call CheckLine(PlayerChar.GetPosR, PlayerChar.GetPosC, PlayerChar.GetPosR - 3, PlayerChar.GetPosC + 3)

End Sub

Sub CheckLine(r1, c1, r2, c2)

Dim dR, dC, D As Double
Dim r, c As Integer
Dim quarter As String

dC = c2 - c1
dR = r2 - r1

If dR = 0 Then
    dR = 1E-20
ElseIf dC = 0 Then
    dC = 1E-20
End If

If c2 > c1 And Abs(dC / dR) >= 1 Then
    quarter = "right"
ElseIf c1 > c2 And Abs(dC / dR) >= 1 Then
    quarter = "left"
ElseIf r2 > r1 And Abs(dC / dR) < 1 Then
    quarter = "down"
ElseIf r1 > r2 And Abs(dC / dR) < 1 Then
    quarter = "up"
End If

Select Case quarter
    Case "right"
        For c = c1 To c2
            r = Round(r1 + dR * (c - c1) / dC)
            If MapOutLine(r, c) < 1 Then
                MapOutLine(r, c) = 0
                Exit Sub
            End If
            MapOutLine(r, c) = 3
        Next
    Case "left"
        For c = c1 To c2 Step -1
            r = Round(r1 + dR * (c - c1) / dC)
            If MapOutLine(r, c) < 1 Then
                MapOutLine(r, c) = 0
                Exit Sub
            End If
            MapOutLine(r, c) = 3
        Next
    Case "down"
        For r = r1 To r2
            c = Round(c1 + dC * (r - r1) / dR)
            If MapOutLine(r, c) < 1 Then
                MapOutLine(r, c) = 0
                Exit Sub
            End If
            MapOutLine(r, c) = 3
        Next
    Case "up"
        For r = r1 To r2 Step -1
            c = Round(c1 + dC * (r - r1) / dR)
            If MapOutLine(r, c) < 1 Then
                MapOutLine(r, c) = 0
                Exit Sub
            End If
            MapOutLine(r, c) = 3
        Next
End Select
End Sub

Sub ReDrawCurrMap(RStart, CStart, REnd, CEnd)

If RStart < 2 Then RStart = 2
If CStart < 2 Then CStart = 2
If REnd > 30 Then REnd = 30
If CEnd > 55 Then CEnd = 55

For r = RStart To REnd
    For c = CStart To CEnd
        Select Case MapOutLine(r, c)
            Case -1
                ICSRH.Cells(r, c).Interior.Color = vbBlack
            Case 0
                ICSRH.Cells(r, c).Interior.Color = RGB(40, 40, 40)
            Case 1
                ICSRH.Cells(r, c).Interior.Color = vbBlack
            Case 2
                ICSRH.Cells(r, c).Interior.Color = RGB(200, 200, 200)
            Case 3
                ICSRH.Cells(r, c).Interior.ColorIndex = xlColorIndexNone
        End Select
    Next
Next

End Sub
