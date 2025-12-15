VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private colLivroID As Collection
Private colLivroNome As Collection
Private colLivroAutor As Collection
Private colLivroEdicao As Collection
Private colLivroDTPubli As Collection
Private GuardaIndiceLivroEncontrado As Integer
Private Encontrado As Boolean
Private msg As String

Public Property Get gLivroID() As Collection
   Set gLivroID = colLivroID
End Property

Public Property Get gLivroNome() As Collection
   Set gLivroNome = colLivroNome
End Property

Public Property Get gLivroAutor() As Collection
   Set gLivroAutor = colLivroAutor
End Property

Public Property Get gLivroEdicao() As Collection
   Set gLivroEdicao = colLivroEdicao
End Property

Public Property Get gLivroDTPubli() As Collection
   Set gLivroDTPubli = colLivroDTPubli
End Property

Public Property Get gmsg() As String
   gmsg = msg
End Property

Public Property Get gGuardaIndiceLivroEncontrado() As Integer
   gGuardaIndiceLivroEncontrado = GuardaIndiceLivroEncontrado
End Property

Public Property Get gEncontrado() As Boolean
   gEncontrado = Encontrado
End Property

Public Function Inserir(strNome As String, intEdicao As Integer, strAutor As String, dtDataPubli As Date) As Boolean

   On Error GoTo Inserir_E
   
   Inserir = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans

   Dim strInsert As String
   strInsert = "INSERT INTO LIVROS_BRAYAN (NOME, EDICAO, DT_PUBLI, AUTOR) VALUES " & "('" & strNome & "', " & intEdicao & ", '" & dtDataPubli & "', '" & strAutor & "')"
   conn.Execute strInsert

   conn.CommitTrans
   conn.Close

   Inserir = True

   GoTo DestruirObjetos

Inserir_E:

   MsgBox "Ocorreu um erro na função Inserir - clsBD, ao adicionar um novo livro ao banco de dados."

   conn.RollBack

   GoTo DestruirObjetos

DestruirObjetos:
   ' Limpar objetos
End Function

Public Function AtualizarRegistro(ID As Integer, Nome As String, Autor As String, Edicao As String, DT_Publi As String) As Boolean

   On Error GoTo AtualizarRegistro_E
   
   AtualizarRegistro = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans

   Dim strUpdate As String
   strUpdate = "UPDATE LIVROS_BRAYAN SET NOME = '" & Nome & "', AUTOR = '" & Autor & "', EDICAO = " & Edicao & ", DT_PUBLI = '" & DT_Publi & "' WHERE ID = " & ID & ";"

   conn.Execute strUpdate

   conn.CommitTrans
   conn.Close

   AtualizarRegistro = True

   GoTo DestruirObjetos

AtualizarRegistro_E:

   MsgBox "Ocorreu um erro na função Inserir - clsBD, ao adicionar um novo livro ao banco de dados."

   conn.RollBack

   GoTo DestruirObjetos

DestruirObjetos:
   ' Limpar objetos
End Function

Public Function Selecionar() As Boolean

   On Error GoTo Selecionar_E
   
   Selecionar = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans

   rs.Open "SELECT * FROM LIVROS_BRAYAN;", conn, adOpenKeyset

   While Not rs.EOF
      colLivroID.Add rs.Fields.Item(0).Value
      colLivroNome.Add rs.Fields.Item(1).Value
      colLivroAutor.Add rs.Fields.Item(2).Value
      colLivroDTPubli.Add rs.Fields.Item(3).Value
      colLivroEdicao.Add rs.Fields.Item(4).Value
      rs.MoveNext
   Wend

   conn.CommitTrans

   rs.Close
   conn.Close

   Selecionar = True

   GoTo DestruirObjetos

Selecionar_E:

   MsgBox "Ocorreu um erro na função Selecionar - clsBD, ao adicionar um novo livro ao banco de dados."

   conn.RollBack

   GoTo DestruirObjetos

DestruirObjetos:
   ' Limpar objetos
End Function

Public Function Pesquisar(listLivros As ListView, termoPesquisa As String) As Boolean

   On Error GoTo Pesquisar_E

   Pesquisar = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans

   rs.Open "SELECT * FROM LIVROS_BRAYAN;", conn, adOpenKeyset

   While Not rs.EOF
      colLivroID.Add rs.Fields.Item(0).Value
      colLivroNome.Add rs.Fields.Item(1).Value
      colLivroAutor.Add rs.Fields.Item(2).Value
      colLivroDTPubli.Add rs.Fields.Item(3).Value
      colLivroEdicao.Add rs.Fields.Item(4).Value
      rs.MoveNext
   Wend
   
   conn.CommitTrans

   rs.Close
   conn.Close

   Dim resultado As String
   For i = 1 To listLivros.ListItems.Count
      If colLivroNome.Item(i) = termoPesquisa Or colLivroAutor.Item(i) = termoPesquisa Or colLivroEdicao.Item(i) = termoPesquisa Or colLivroDTPubli.Item(i) = termoPesquisa Then
        resultado = "O livro foi encontrado foi " & colLivroNome.Item(i) & ", " & colLivroAutor.Item(i) & ", " & "Edição " & colLivroEdicao.Item(i) & ", " & "Data de publicação: " & colLivroDTPubli.Item(i) & "!"
        GuardaIndiceLivroEncontrado = i
        Encontrado = True
        Exit For
     End If
   Next i

   ' Exibe o resultado na caixa de texto
   If Encontrado Then
      msg = resultado
   Else
      msg = "Nenhum livro encontrado!"
   End If

   Pesquisar = True

   GoTo DestruirObjetos

Pesquisar_E:
DestruirObjetos:
End Function

Public Function PesquisarComListView(listLivros As ListView) As Boolean

   On Error GoTo PesquisarListView_E

   PesquisarListView = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans

   rs.Open "SELECT * FROM LIVROS_BRAYAN;", conn, adOpenKeyset

   While Not rs.EOF
      colLivroID.Add rs.Fields.Item(0).Value
      colLivroNome.Add rs.Fields.Item(1).Value
      colLivroAutor.Add rs.Fields.Item(2).Value
      colLivroDTPubli.Add rs.Fields.Item(3).Value
      colLivroEdicao.Add rs.Fields.Item(4).Value
      rs.MoveNext
   Wend
   
   conn.CommitTrans

   rs.Close
   conn.Close

   Dim resultado As String
   For i = 1 To listLivros.ListItems.Count
      If colLivroID(i) = listLivros.SelectedItem Then
        resultado = "O livro foi encontrado foi " & colLivroNome.Item(i) & ", " & colLivroAutor.Item(i) & ", " & "Edição " & colLivroEdicao.Item(i) & ", " & "Data de publicação: " & colLivroDTPubli.Item(i) & "!"
        GuardaIndiceLivroEncontrado = i
        Encontrado = True
        Exit For
     End If
   Next i

   ' Exibe o resultado na caixa de texto
   If Encontrado Then
      msg = resultado
   Else
      msg = "Nenhum livro encontrado!"
   End If

   PesquisarListView = True

   GoTo DestruirObjetos

PesquisarListView_E:
DestruirObjetos:
End Function

'Public Function Excluir(listLivros As ListView) As Boolean
'
'   On Error GoTo Excluir_E
'
'   Excluir = False
'
'   Dim conn As New ADODB.Connection
'   Dim rs As New ADODB.Recordset
'   Dim connString As String
'
'   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
'   conn.Open connString
'
'   conn.BeginTrans
'
'   rs.Open "SELECT * FROM LIVROS_BRAYAN;", conn, adOpenKeyset
'
'   While Not rs.EOF
'      colLivroID.Add rs.Fields.Item(0).Value
'      colLivroNome.Add rs.Fields.Item(1).Value
'      colLivroAutor.Add rs.Fields.Item(2).Value
'      colLivroDTPubli.Add rs.Fields.Item(3).Value
'      colLivroEdicao.Add rs.Fields.Item(4).Value
'      rs.MoveNext
'   Wend
'
'   rs.Close
'
'   For i = 1 To listLivros.ListItems.Count
'      ID = colLivroID(i)
'      If listLivros.ListItems(i).Selected Then
'         Dim strDelete As String
'         strDelete = "DELETE FROM LIVROS_BRAYAN WHERE ID = " & ID
'         conn.Execute strDelete
'         msg = "Livro " & colLivroNome(i) & " excluido!"
'         Exit For
'      End If
'   Next i
'
'   conn.CommitTrans
'
'   conn.Close
'
'   Excluir = True
'
'   GoTo DestruirObjetos
'
'Excluir_E:
'DestruirObjetos:
'End Function

Public Function Excluir() As Boolean

   On Error GoTo Excluir_E

   Excluir = False

   Dim conn As New ADODB.Connection
   Dim rs As New ADODB.Recordset
   Dim connString As String

   connString = "Driver=SQL Server;Server=10.11.12.7\lydsqlhml;Database=Desenv1;User Id=lydians_mdc;Password=lydians_mdc"
   conn.Open connString

   conn.BeginTrans
   
   ID = colLivroID(GuardaIndiceLivroEncontrado)
   Dim strDelete As String
   strDelete = "DELETE FROM LIVROS_BRAYAN WHERE ID = " & ID
   conn.Execute strDelete
   msg = "Livro " & colLivroNome(GuardaIndiceLivroEncontrado) & " excluido!"
   
   conn.CommitTrans

   conn.Close

   Excluir = True

   GoTo DestruirObjetos

Excluir_E:
DestruirObjetos:
End Function

Private Sub Class_Initialize()
   Set colLivroID = New Collection
   Set colLivroAutor = New Collection
   Set colLivroDTPubli = New Collection
   Set colLivroEdicao = New Collection
   Set colLivroNome = New Collection
End Sub

Private Sub Class_Terminate()
   Set colLivroAutor = Nothing
   Set colLivroDTPubli = Nothing
   Set colLivroEdicao = Nothing
   Set colLivroNome = Nothing
End Sub
