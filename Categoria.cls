VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Categoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private v_nome  As String
Private v_valor As Double
Private v_id As Integer


Public Property Get Nome() As String
    Nome = v_nome
End Property

Public Property Let Nome(ByVal vNewValue As String)
    v_nome = vNewValue
End Property


Public Property Get Valor() As Double
    Valor = v_valor
End Property

Public Property Let Valor(ByVal vNewValue As Double)
    v_valor = vNewValue
End Property

Public Property Get Id() As Integer
    Id = v_id
    

End Property

Public Property Let Id(ByVal vNewValue As Integer)
    v_id = vNewValue
End Property
