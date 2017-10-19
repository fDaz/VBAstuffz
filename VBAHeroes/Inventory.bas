Attribute VB_Name = "Inventory"
Private InvList As New Collection

Sub Init()

End Sub

Function GetInvSize()

GetInvSize = InvList.Count

End Function

Sub AddToInventory(ItemObject As Object)

Select Case TypeName(ItemObject)
    Case "cWeapon", "cArmor"
        InvList.Add ItemObject
End Select

End Sub

Sub RemoveFromInventory(InvPos)

If InvPos > InvList.Count Then
    MsgBox "No item on player with the position: " & InvPos
    Exit Sub
End If

InvList.Remove (InvPos)

End Sub

Function ReturnItemByInvPos(InvPos) As Object

If InvPos > InvList.Count Then
    MsgBox "No item on player with the position: " & InvPos
    Exit Function
Else
    ReturnItemByInvPos = ItemList(InvPos)
End If

End Function

Sub ShowInventory()

Call Windws.InitWindow(4, 18, 28, 39)

ICSRH.Cells(6, 20) = "Inventory:"

Call ListInventory

Call ICSRH.SetControlType(1)

End Sub

Sub ShowDropMenu()

Call Windws.InitWindow(4, 18, 28, 39)

ICSRH.Cells(6, 20) = "Select an item to drop:"

Call ListInventory

Call ICSRH.SetControlType(2)

End Sub

Sub ShowAppMenu()

Call Windws.InitWindow(4, 18, 28, 39)

ICSRH.Cells(6, 20) = "Select an item to appraise:"

Call ListInventory

Call ICSRH.SetControlType(4)

End Sub

Sub ListInventory()

If InvList.Count = 0 Then
    ICSRH.Cells(8, 20) = "Nothing"
Else
    For i = 1 To InvList.Count
        ICSRH.Cells(7 + i, 20) = GetKey(i) & ")"
        ICSRH.Cells(7 + i, 21) = InvList(i).Name
    Next
End If

    ICSRH.Cells(26, 20) = "z)"
    ICSRH.Cells(26, 21) = "Exit"

End Sub

Sub UseInvItem(InvPos)

Select Case TypeName(InvList(InvPos)) 'check what kind of item is it in the first place
    Case "cWeapon" 'wield
        If TypeName(PlayerChar.GetEq("W")) = "cWeapon" Then
            Call AddToInventory(PlayerChar.GetEq("W"))
        End If
        Call PlayerChar.SetEq("W", InvList(InvPos))
        Call MessageLog.NewMessage("You've got the " & Format(InvList(InvPos).Name, "<") & " at the ready")
        Call RemoveFromInventory(InvPos)
        Call PlayerChar.DrawStats
    Case "cArmor" 'armor/helm/boots equip
        Select Case InvList(InvPos).Typ
            Case "body"
                If TypeName(PlayerChar.GetEq("A")) = "cArmor" Then
                    Call AddToInventory(PlayerChar.GetEq("A"))
                End If
                Call PlayerChar.SetEq("A", InvList(InvPos))
            Case "head"
                If TypeName(PlayerChar.GetEq("H")) = "cArmor" Then
                    Call AddToInventory(PlayerChar.GetEq("H"))
                End If
                Call PlayerChar.SetEq("H", InvList(InvPos))
            Case "feet"
                If TypeName(PlayerChar.GetEq("B")) = "cArmor" Then
                    Call AddToInventory(PlayerChar.GetEq("B"))
                End If
                Call PlayerChar.SetEq("B", InvList(InvPos))
            End Select
        Call MessageLog.NewMessage("You put on the " & Format(InvList(InvPos).Name, "<"))
        Call RemoveFromInventory(InvPos)
        Call PlayerChar.DrawStats
    End Select
ICSRH.IncRounds
Call Windws.CloseWindow(4, 14, 28, 43)

End Sub

Sub DropItem(InvPos)

Call MessageLog.NewMessage("Dropped: " & InvList(InvPos).Name)
InvList(InvPos).PosR = PlayerChar.GetPosR
InvList(InvPos).PosC = PlayerChar.GetPosC
Call DepthMap.PlaceItem(InvList(InvPos))
Call RemoveFromInventory(InvPos)

ICSRH.IncRounds
Call Windws.CloseWindow(4, 14, 28, 43)
                    
End Sub

Sub AppItem(InvPos)

Select Case TypeName(InvList(InvPos))
    Case "cWeapon"
        Call MessageLog.NewMessage("Its base damage potential is " & InvList(InvPos).QualityDesc)
        If InvList(InvPos).StrScaling > InvList(InvPos).DexScaling Then
            Call MessageLog.NewMessage("It scales better with strength")
        ElseIf InvList(InvPos).StrScaling < InvList(InvPos).DexScaling Then
            Call MessageLog.NewMessage("It scales better with dexterity")
        End If
    Case "cArmor"
        Call MessageLog.NewMessage("Its base defense potential is " & InvList(InvPos).QualityDesc(0))
        Select Case InvList(InvPos).QualityDesc(1)
            Case "light"
                Call MessageLog.NewMessage("It doesn't seem like it would hamper you much")
            Case "medium"
                Call MessageLog.NewMessage("It seem like it would hamper you a bit")
            Case "heavy"
                Call MessageLog.NewMessage("It seem like it would hamper you a fair amount")
        End Select
End Select

ICSRH.IncRounds
Call Windws.CloseWindow(4, 14, 28, 43)

End Sub
