VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PositionR, PositionC, MaterialID, ItemtypeID As Integer

Property Get PosR() As Integer
    PosR = PositionR
End Property

Property Get PosC() As Integer
    PosC = PositionC
End Property

Property Let PosR(value As Integer)
    PositionR = value
End Property

Property Let PosC(value As Integer)
    PositionC = value
End Property

Property Let MatID(value As Integer)
    MaterialID = value
End Property

Property Let ItemID(value As Integer)
    ItemtypeID = value
End Property

Property Get Typ() As String
    Typ = ItemList.ItemAttr("cWeapon", ItemtypeID, 3)
End Property

Property Get Atk() As Integer
    Select Case Typ
        Case "blunt"
            Atk = ItemList.MatAttr(MaterialID, 7) + ItemList.ItemAttr("cWeapon", ItemtypeID, 5)
        Case "edge", "pierce"
            Atk = ItemList.MatAttr(MaterialID, 6) + ItemList.ItemAttr("cWeapon", ItemtypeID, 5)
    End Select
End Property

Property Get Name() As String
    Name = ItemList.MatAttr(MaterialID, 1) & " " & ItemList.ItemAttr("cWeapon", ItemtypeID, 1)
End Property

Property Get Icon() As String
    Icon = ItemList.ItemAttr("cWeapon", ItemtypeID, 2)
End Property

Property Get Color(which As String) As Integer
    Select Case which
        Case "R"
            Color = ItemList.MatAttr(MaterialID, 2)
        Case "G"
            Color = ItemList.MatAttr(MaterialID, 3)
        Case "B"
            Color = ItemList.MatAttr(MaterialID, 4)
    End Select
End Property

Property Get StrScaling() As Integer
    StrScaling = ItemList.ItemAttr("cWeapon", ItemtypeID, 6)
End Property

Property Get DexScaling() As Integer
    DexScaling = ItemList.ItemAttr("cWeapon", ItemtypeID, 7)
End Property

Property Get QualityDesc() As String
    Select Case Typ
        Case "blunt"
            MinDens = ItemList.GetMinVal(0, 7)
            MaxDens = ItemList.GetMaxVal(0, 7)
            
            MinAtk = ItemList.GetMinVal(1, 5)
            MaxAtk = ItemList.GetMaxVal(1, 5)
            
            If Atk < (((MaxDens + MaxAtk) - (MinDens + MinAtk)) / 3) + (MinDens + MinAtk) Then
                QualityDesc = "poor"
            ElseIf Atk < (((MaxDens + MaxAtk) - (MinDens + MinAtk)) / 3) * 2 + (MinDens + MinAtk) Then
                QualityDesc = "average"
            Else
                QualityDesc = "good"
            End If
        Case "edge", "pierce"
            MinHard = ItemList.GetMinVal(0, 6)
            MaxHard = ItemList.GetMaxVal(0, 6)
            
            MinAtk = ItemList.GetMinVal(1, 5)
            MaxAtk = ItemList.GetMaxVal(1, 5)
            
            If Atk < (((MaxHard + MaxAtk) - (MinHard + MinAtk)) / 3) + (MinHard + MinAtk) Then
                QualityDesc = "poor"
            ElseIf Atk < (((MaxHard + MaxAtk) - (MinHard + MinAtk)) / 3) * 2 + (MinHard + MinAtk) Then
                QualityDesc = "average"
            Else
                QualityDesc = "good"
            End If
    End Select
End Property
