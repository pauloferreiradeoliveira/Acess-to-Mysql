Attribute VB_Name = "SubCategoria"
Option Compare Database

Dim conexaoMySql As New conexaoMySql
Dim conexaoAccess As New conexaoAccess

Sub subCategoria()

    Dim rstMySql As ADODB.Recordset
    Dim rstAccess As DAO.Recordset
    
    
    Set rstAccess = conexaoAccess.getRecord("SELECT  SubCategoria.nome,SubCategoria.PK_Categoria,Categoria.Nome as Categoria FROM subCategoria INNER JOIN Categoria ON Categoria.id_categoria = SubCategoria.PK_Categoria")
    rstAccess.MoveFirst
    While Not rstAccess.EOF
        Dim sql As String
        sql = "select * from Categoria where Categoria.nome = '" & rstAccess!Categoria & "'"
        Set rstMySql = conexaoMySql.consultaSQL(sql)
        rstMySql.MoveFirst
        
        sql = "insert into subCategoria(nome,nomeCategoria,fk_Categoria) values ('" & rstAccess!Nome & "','" & rstAccess!Categoria & "'," & rstMySql!idCategoria & ")"
        MsgBox (sql)
        conexaoMySql.consultaSQL (sql)
        rstAccess.MoveNext
    Wend
        
End Sub
