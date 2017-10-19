VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cArmor"
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
    Typ = ItemList.ItemAttr("cArmor", ItemtypeID, 3)
End Property

Property Get Def() As Integer
    Def = ItemList.MatAttr(MaterialID, 6) + ItemList.ItemAttr("cArmor", ItemtypeID, 5)
End Property

Property Get Dens() As Integer
    Dens = ItemList.MatAttr(MaterialID, 7)
End Property

Property Get Name() As String
    Name = ItemList.MatAttr(MaterialID, 1) & " " & ItemList.ItemAttr("cArmor", ItemtypeID, 1)
End Property

Property Get Icon() As String
    Icon = ItemList.ItemAttr("cArmor", ItemtypeID, 2)
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

Property Get QualityDesc(what As Integer) As String 'what: 0 defense, 1 dodge penalty
    Select Case what
        Case 0
            MinHard = ItemList.GetMinVal(0, 6)
            MaxHard = ItemList.GetMaxVal(0, 6)
            
            MinDef = ItemList.GetMinVal(2, 5)
            MaxDef = ItemList.GetMaxVal(2, 5)
            
            If Def < (((MaxHard + MaxDef) - (MinHard + MinDef)) / 3) + (MinHard + MinDef) Then
                QualityDesc = "poor"
            ElseIf Def < (((MaxHard + MaxDef) - (MinHard + MinDef)) / 3) * 2 + (MinHard + MinDef) Then
                QualityDesc = "average"
            Else
                QualityDesc = "good"
            End If
        Case 1
            MinDens = ItemList.GetMinVal(0, 7)
            MaxDens = ItemList.GetMaxVal(0, 7)
            
            If Dens < ((MaxDens - MinDens) / 3) + MinDens Then
                QualityDesc = "light"
            ElseIf Dens < ((MaxDens - MinDens) / 3) * 2 + MinDens Then
                QualityDesc = "medium"
            Else
                QualityDesc = "heavy"
            End If
    End Select
End Property
