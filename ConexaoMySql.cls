VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConexaoMySql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private valorConexao As String
Private conexao As ADODB.Connection

Private Function pegarConexao() As ADODB.Connection
    valorConexao = "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; DATABASE=test1; UID=root; PASSWORD=root; option=3"
    If conexao Is Nothing Then
        Set conexao = New ADODB.Connection
        conexao.Open (valorConexao)
    End If
End Function

Public Function fecharConexao()
    conexao.Close
End Function

Public Function consultaSQL(ByVal newSql As String) As ADODB.Recordset
    pegarConexao
    Set consultaSQL = conexao.Execute(newSql)
End Function
