Attribute VB_Name = "ItemList"
Private WeaponLst As New Collection
Private ArmorLst As New Collection
Private MaterialLst As New Collection

Sub Init()

Dim i As Integer
Dim arrayToDict(15)

'Get the materials

i = 2 'first actual row of item table

Do

arrayToDict(0) = i - 1

For n = 1 To UBound(arrayToDict)
    arrayToDict(n) = Materials.Cells(i, n)
Next

MaterialLst.Add arrayToDict

i = i + 1

Loop Until Materials.Cells(i, 1) = ""

'Plan: 3 sheets: weapon, armor, other;
'Weapons

i = 2 'first actual row of item table

Do

arrayToDict(0) = i - 1

For n = 1 To UBound(arrayToDict)
    arrayToDict(n) = WpnSheet.Cells(i, n)
Next

WeaponLst.Add arrayToDict

i = i + 1

Loop Until WpnSheet.Cells(i, 1) = ""

'Armors

i = 2 'first actual row of item table

Do

arrayToDict(0) = i - 1

For n = 1 To UBound(arrayToDict)
    arrayToDict(n) = ArmrSheet.Cells(i, n)
Next

ArmorLst.Add arrayToDict

i = i + 1

Loop Until ArmrSheet.Cells(i, 1) = ""

End Sub

Function ItemAttr(what, ItemID, attr)

Select Case what
    Case "cWeapon"
        ItemAttr = WeaponLst(ItemID)(attr)
    Case "cArmor"
        ItemAttr = ArmorLst(ItemID)(attr)
End Select

End Function

Function MatAttr(MaterialID, attr)

MatAttr = MaterialLst(MaterialID)(attr)

End Function

Function ValidItem(ItemID)

ValidItem = ItemLst.Exists(ItemID)

End Function

Function ValidMat(MatID)

ValidMat = MaterialLst.Exists(MatID)

End Function

Function ListLength(what) '0: material, 1: weapons, 2: armors

Select Case what
    Case 0
        ListLength = MaterialLst.Count
    Case 1
        ListLength = WeaponLst.Count
    Case 2
        ListLength = ArmorLst.Count
End Select
End Function

Function GetMinVal(list, index)

Dim RetVal As Integer

RetVal = 999

Select Case list
    Case 0
        For i = 1 To MaterialLst.Count
            If MaterialLst(i)(index) < RetVal Then RetVal = MaterialLst(i)(index)
        Next
    Case 1
        For i = 1 To WeaponLst.Count
            If WeaponLst(i)(index) < RetVal Then RetVal = WeaponLst(i)(index)
        Next
    Case 2
        For i = 1 To ArmorLst.Count
            If ArmorLst(i)(index) < RetVal Then RetVal = ArmorLst(i)(index)
        Next
End Select

GetMinVal = RetVal

End Function

Function GetMaxVal(list, index)

Dim RetVal As Integer

RetVal = 0

Select Case list
    Case 0
        For i = 1 To MaterialLst.Count
            If MaterialLst(i)(index) > RetVal Then RetVal = MaterialLst(i)(index)
        Next
    Case 1
        For i = 1 To WeaponLst.Count
            If WeaponLst(i)(index) > RetVal Then RetVal = WeaponLst(i)(index)
        Next
    Case 2
        For i = 1 To ArmorLst.Count
            If ArmorLst(i)(index) > RetVal Then RetVal = ArmorLst(i)(index)
        Next
End Select

GetMaxVal = RetVal

End Function
