Attribute VB_Name = "Pegando Acess to MySQL"
Option Compare Database
Sub testSelect()
    
    'Cria uma conexção
    Dim dbs As DAO.Database
    'Para poder ler
    Dim rstCategorias As DAO.Recordset
    'SQL para ser realizado
    Dim strSQL As String
    
    'Parte MSQL
    Dim mysConString As String
    'Verifica o ODBC
    mysConString = "DRIVER={MySQL ODBC 8.0 ANSI Driver}; SERVER=localhost; DATABASE=test1; UID=root; PASSWORD=root; OPTION=3;"
    'Verifica de o 'Microsoft ActiveX Data Objects x. x ' da ativo
    Dim con As New ADODB.Connection
    con.Open (mysConString)
    
    'Caregando o DB Local
    Set dbs = CurrentDb
    'Criando o SQL
    strSQL = "SELECT * FROM Categoria"
    Set rstCategorias = dbs.OpenRecordset(strSQL, dbOpenSnapshot)
       
    While Not rstCategorias.EOF
        Dim categoria As New categoria
        categoria.Nome = rstCategorias!Nome
        categoria.Valor = rstCategorias!Valor
        con.Execute ("insert into categoria(Nome) values ('" & categoria.Nome & "')")
        rstCategorias.MoveNext
    Wend
    
    con.Close
    dbs.Close
    
End Sub
