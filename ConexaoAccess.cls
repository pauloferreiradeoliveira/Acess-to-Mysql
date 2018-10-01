VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConexaoAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private conexao As DAO.Database

Public Function criarConexao()
    If conexao Is Nothing Then
        Set conexao = CurrentDb
    End If
End Function

Public Function finalizarConexao()
     If conexao Is Nothing Then
        conexao.Close
    End If
End Function

Public Function getRecord(ByVal vSQL As String) As DAO.Recordset
    criarConexao
    Set getRecord = conexao.OpenRecordset(vSQL, dbOpenSnapshot)
End Function
