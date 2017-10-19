Attribute VB_Name = "PlayerChar"
Private PlPosR, PlPosC As Integer
Private Attribs As New Scripting.Dictionary
Private Stats As New Scripting.Dictionary
Private Equipment As New Scripting.Dictionary
Private Skills As New Scripting.Dictionary
Private DefaultW As New cDefault
Private DefaultA As New cDefault
Private DefaultH As New cDefault
Private DefaultB As New cDefault

Sub Draw()

ICSRH.Cells(PlPosR, PlPosC).Font.Color = vbBlack
ICSRH.Cells(PlPosR, PlPosC) = "@"

End Sub

Sub DrawStats()

Cells(3, 58) = Equipment("W").Name

Cells(3, 60) = Equipment("A").Name

Cells(4, 58) = Equipment("H").Name

Cells(4, 60) = Equipment("B").Name

Cells(5, 58) = Stats("Exp")

Cells(5, 60) = Stats("Lvl")

Cells(6, 58) = Stats("HP") & " / " & Stats("MaxHP")

Cells(6, 60) = Stats("SP") & " / " & Stats("MaxSP")


End Sub

Sub Init()

Attribs.Add "Str", 5
Attribs.Add "Dex", 5
Attribs.Add "End", 5
Attribs.Add "Int", 5
Attribs.Add "Lck", 5

Stats.Add "HP", 15
Stats.Add "MaxHP", 15
Stats.Add "SP", 10
Stats.Add "MaxSP", 10
Stats.Add "Exp", 0
Stats.Add "Lvl", 1
Stats.Add "Atk", 1
Stats.Add "Tohit", 1
Stats.Add "Def", 1
Stats.Add "Dodge", 1
Stats.Add "Crit", 1

DefaultW.Name = "Fists"
DefaultA.Name = "Clothes"
DefaultH.Name = "Nothing"
DefaultB.Name = "Sandals"

Equipment.Add "W", DefaultW
Equipment.Add "A", DefaultA
Equipment.Add "H", DefaultH
Equipment.Add "B", DefaultB

Call RecalcStats
Call DrawStats

End Sub

Sub MovePlayer(dir)

Dim newPlayerPos(1) As Integer

newPlayerPos(0) = PlPosR
newPlayerPos(1) = PlPosC

Select Case dir
    Case 1
        newPlayerPos(0) = PlPosR + 1
        newPlayerPos(1) = PlPosC - 1
    Case 2
        newPlayerPos(0) = PlPosR + 1
    Case 3
        newPlayerPos(0) = PlPosR + 1
        newPlayerPos(1) = PlPosC + 1
    Case 4
        newPlayerPos(1) = PlPosC - 1
    Case 5
        
    Case 6
        newPlayerPos(1) = PlPosC + 1
    Case 7
        newPlayerPos(0) = PlPosR - 1
        newPlayerPos(1) = PlPosC - 1
    Case 8
        newPlayerPos(0) = PlPosR - 1
    Case 9
        newPlayerPos(0) = PlPosR - 1
        newPlayerPos(1) = PlPosC + 1
End Select

'Checks the map matrix in DepthMap

If DepthMap.GetTile(newPlayerPos(0), newPlayerPos(1)) < 1 Then
    Exit Sub
End If

PlPosR = newPlayerPos(0)
PlPosC = newPlayerPos(1)

DepthMap.StuffAtPlayerPos
DepthMap.Refresh

ICSRH.IncRounds

End Sub

Sub InteractK()

If PlPosR = DepthExit.GetPosR And PlPosC = DepthExit.GetPosC Then
    ICSRH.IncRounds
    GenDepth.GenMap
End If

End Sub

Sub SetPosR(r)

PlPosR = r

End Sub

Sub SetPosC(c)

PlPosC = c

End Sub

Function GetPosR()

GetPosR = PlPosR

End Function

Function GetPosC()

GetPosC = PlPosC

End Function

Sub SetStat(what As String, amount As Integer)

Stat(what) = amount

End Sub

Function GetStat(what As String)

GetStat = Stat(what)

End Function

Sub SetEq(where As String, what As Object)

Set Equipment(where) = what

End Sub

Function GetEq(where As String) As Object

Set GetEq = Equipment(where)

End Function

Sub RecalcStats()

Stats("MaxHP") = Int(10 + Attribs("End"))
Stats("MaxSP") = Int(Attribs("End") * 2)

Stats("Atk") = Int(Attribs("Str") + Attribs("Dex") / 2)
Stats("Tohit") = Int(Attribs("Dex") + Attribs("Str") / 2 + Stats("Lvl"))
Stats("Def") = Int(Attribs("End") + Attribs("Dex") / 2)
Stats("Dodge") = Int(Attribs("Dex") + Attribs("End") / 2 + Stats("Lvl"))
Stats("Crit") = Int((2.5 * Attribs("Lck")) / (0.05 * Attribs("Lck") + 1))

End Sub

Sub ShowCharSheet()

Call Windws.InitWindow(4, 18, 28, 39)

ICSRH.Cells(6, 20) = "Stats:"

ICSRH.Cells(8, 20) = "Str:"
ICSRH.Cells(8, 22) = Attribs("Str")

ICSRH.Cells(9, 20) = "Dex:"
ICSRH.Cells(9, 22) = Attribs("Dex")

ICSRH.Cells(10, 20) = "End:"
ICSRH.Cells(10, 22) = Attribs("End")

ICSRH.Cells(11, 20) = "Lck:"
ICSRH.Cells(11, 22) = Attribs("Lck")

ICSRH.Cells(8, 29) = "Tohit:"
ICSRH.Cells(8, 32) = Stats("Tohit")

ICSRH.Cells(9, 29) = "Dodge:"
ICSRH.Cells(9, 32) = Stats("Dodge")

ICSRH.Cells(10, 29) = "Crit:"
ICSRH.Cells(10, 32) = Stats("Crit")

ICSRH.Cells(11, 29) = "To next level: " & Stats("Lvl") * 100 - Stats("Exp") & " Exp"

ICSRH.Cells(13, 20) = "Known skills:"

ICSRH.Cells(26, 20) = "z)"
ICSRH.Cells(26, 21) = "Exit"

Call ICSRH.SetControlType(5)

End Sub
