VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private DName As String
Private dAtr As Integer

Property Get Name() As String
    Name = DName
End Property

Property Let Name(value As String)
    DName = value
End Property

Property Get Atr() As Integer
    Atr = dAtr
End Property

Property Let PosC(value As Integer)
    dAtr = value
End Property
