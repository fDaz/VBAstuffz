Attribute VB_Name = "GenDepth"
Private PrevCntrR, PrevCntrC, FirstRoom(1), RoomNum, MaxRooms, UpLimit, LowLimit As Integer

Sub GenMap()

'Resets Game Area
ICSRH.Range(Cells(2, 2), Cells(30, 55)).Interior.Color = vbBlack
ICSRH.Range(Cells(2, 2), Cells(30, 55)) = ""
Call DepthMap.Init

ICSRH.IncDepth
Cells(2, 58) = ICSRH.GetDepth

RoomNum = 0

UpLimit = 8
LowLimit = 5

Randomize
MaxRooms = Int((UpLimit - LowLimit + 1) * Rnd + LowLimit)

If ICSRH.GetDepth = 1 Then
    PrevCntrR = 16
    PrevCntrC = 28
    FirstRoom(0) = 16
    FirstRoom(1) = 28
End If

Do While RoomNum < MaxRooms
    RoomNum = RoomNum + 1
    Call GenRoom
Loop

Call ConnectRoom(FirstRoom(0), FirstRoom(1))

DepthMap.SaveCurrMap

ICSRH.Range(Cells(2, 2), Cells(30, 55)) = ""

DepthMap.Refresh

ICSRH.EnterComm.Activate

End Sub

Private Sub GenRoom()

Dim RRad, CRad, Center(1), ItemPlace(1), ItemID, MaterialID As Integer

Dim DBug As Integer 'If bigger than 100 won't keep trying to gen a new room

DBug = 0

Do

DBug = DBug + 1

'From 2nd level onwards, use the last depth's last room for the first room of the new level
If ICSRH.GetDepth > 1 And RoomNum = 1 Then

RRad = 1
CRad = 2

Center(0) = PrevCntrR
Center(1) = PrevCntrC
FirstRoom(0) = PrevCntrR
FirstRoom(1) = PrevCntrC

Else

'Generate the room size
Randomize
RRad = 1 + Int((Rnd * 4))
Randomize
CRad = 2 + Int((Rnd * 4))

'Generate the room position: Center(0) is the row, Center (1) is the column
Randomize
Center(0) = Int(((29 - RRad) - (3 + RRad) + 1) * Rnd + (3 + RRad))
Randomize
Center(1) = Int(((54 - CRad) - (3 + CRad) + 1) * Rnd + (3 + CRad))

End If

If DBug > 100 Then
'    MsgBox "Can't place more rooms!"
    MaxRooms = RoomNum
    DepthExit.SetPosR (PrevCntrR)
    DepthExit.SetPosC (PrevCntrC)
    Exit Sub
End If

Loop While IsThereSpace(Center(0) - (RRad + 2), Center(1) - (CRad + 2), Center(0) + (RRad + 2), Center(1) + (CRad + 2)) = False

'Draw The Room

ICSRH.Range(Cells(Center(0) - RRad, Center(1) - CRad), Cells(Center(0) + RRad, Center(1) + CRad)) = "RM"

'Place the goodies: ItemPlace(0) is the row, ItemPlace(1) is the column
'First place weapons
Randomize
If Rnd * 100 < 70 Then
    Randomize
    ItemPlace(0) = Int(((Center(0) + RRad) - (Center(0) - RRad) + 1) * Rnd + (Center(0) - RRad))
    Randomize
    ItemPlace(1) = Int(((Center(1) + CRad) - (Center(1) - CRad) + 1) * Rnd + (Center(1) - CRad))
    'Decide what to place - the more you play, the better stuff you get
    Do
        Randomize
        ItemID = (Int(Rnd * ItemList.ListLength(1)) + 1)
    Loop Until ItemList.ItemAttr("cWeapon", ItemID, 4) <= Int((ICSRH.GetDepth / 5) + 1)
    Do
        Randomize
        MaterialID = (Int(Rnd * ItemList.ListLength(0)) + 1)
    Loop Until ItemList.MatAttr(MaterialID, 5) <= Int((ICSRH.GetDepth / 5) + 1) And ItemList.MatAttr(MaterialID, 8) = 1
    'All set - place the item on the map matrix
    Call DepthMap.PlaceWeapon(ItemPlace(0), ItemPlace(1), ItemID, MaterialID)
End If

'Then place armor
Randomize
If Rnd * 100 < 70 Then
    Randomize
    ItemPlace(0) = Int(((Center(0) + RRad) - (Center(0) - RRad) + 1) * Rnd + (Center(0) - RRad))
    Randomize
    ItemPlace(1) = Int(((Center(1) + CRad) - (Center(1) - CRad) + 1) * Rnd + (Center(1) - CRad))
    'Decide what to place - the more you play, the better stuff you get
    Do
        Randomize
        ItemID = (Int(Rnd * ItemList.ListLength(2)) + 1)
    Loop Until ItemList.ItemAttr("cArmor", ItemID, 4) <= Int((ICSRH.GetDepth / 5) + 1)
    Do
        Randomize
        MaterialID = (Int(Rnd * ItemList.ListLength(0)) + 1)
    Loop Until ItemList.MatAttr(MaterialID, 5) <= Int((ICSRH.GetDepth / 5) + 1) And ItemList.MatAttr(MaterialID, 9) = 1
    'All set - place the item on the map matrix
    Call DepthMap.PlaceArmor(ItemPlace(0), ItemPlace(1), ItemID, MaterialID)
End If

'Connect Room to the previous one
Call ConnectRoom(Center(0), Center(1))

'Place the player if it's the first room
If RoomNum = 1 Then
    PlayerChar.SetPosR (Center(0))
    PlayerChar.SetPosC (Center(1))
End If

'Place the exit if it's the last room
If RoomNum = MaxRooms Then
    DepthExit.SetPosR (Center(0))
    DepthExit.SetPosC (Center(1))
End If

PrevCntrR = Center(0)
PrevCntrC = Center(1)

End Sub

Private Function IsThereSpace(startR, startC, endR, endC)

If startR < 3 Then startR = 3
If startC < 3 Then startC = 3

If endR > 29 Then endR = 29
If endC > 54 Then endC = 54

For c = startC To endC
    For r = startR To endR
        If Cells(r, c) = "RM" Then
            IsThereSpace = False
            Exit Function
        End If
    Next
Next

IsThereSpace = True

End Function

Private Sub ConnectRoom(r, c)

ICSRH.Range(Cells(r, c), Cells(PrevCntrR, c)) = "RM"
ICSRH.Range(Cells(PrevCntrR, c), Cells(PrevCntrR, PrevCntrC)) = "RM"

End Sub
