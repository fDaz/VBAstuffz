Attribute VB_Name = "Controls"
Private ControlArray(1, 41) As Variant
Private keyID As Integer

Sub Init()

For i = 1 To UBound(ControlArray, 2)
    ControlArray(0, i) = Format(CtrlSheet.Cells(i, 1))
    ControlArray(1, i) = Format(CtrlSheet.Cells(i, 2))
Next

End Sub

Function GetKey(which)

GetKey = ControlArray(0, which)

End Function

Function FindID(keypressed)

For i = 1 To UBound(ControlArray, 2)
    If ControlArray(0, i) = keypressed Then
        FindID = i
        Exit Function
    End If
Next

MsgBox "No key found!"
FindID = 0

End Function

Function GetAction(action)

For i = 1 To UBound(ControlArray)
    If ControlArray(1, i) = action Then
        GetAction = ControlAray(0, i)
        Exit Function
    End If
Next

MsgBox "No action found!"
GetAction = 0

End Function

Sub MapInpt(inpt)

keyID = FindID(inpt)

If keyID >= 28 And keyID <= 36 Then
    Call PlayerChar.MovePlayer(inpt)
    Exit Sub
End If

Select Case ControlArray(1, keyID)
    Case "inventory"
        Call Inventory.ShowInventory
    Case "use"
        Call PlayerChar.InteractK
    Case "drop"
        Call Inventory.ShowDropMenu
    Case "get"
        Call DepthMap.PickUpItem
    Case "help"
        Call MessageLog.NewMessage("No help yet!")
    Case "appraise"
        Call Inventory.ShowAppMenu
    Case "character"
        Call PlayerChar.ShowCharSheet
    Case Else
        Exit Sub
End Select
    
End Sub

Sub MenuInpt(inpt)

keyID = FindID(inpt)

Select Case keyID
    Case 1 To 15
        Select Case ICSRH.GetControlType '0: map control, 1: inventory (use), 2:drop, 3: pick up, 4: appraise, 5:char stats
            Case 1
                If Inventory.GetInvSize >= keyID Then Call Inventory.UseInvItem(keyID)
            Case 2
                If Inventory.GetInvSize >= keyID Then Call Inventory.DropItem(keyID)
            Case 3
                Call DepthMap.SelectWhichItemToPickUp(keyID)
            Case 4
                If Inventory.GetInvSize >= keyID Then Call Inventory.AppItem(keyID)
        End Select
    Case 26
        Call Windws.CloseWindow(4, 14, 28, 43)
End Select

End Sub
