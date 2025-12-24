VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImportadorBasesCorreio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private conexaoComBancoDeDados As ADODB.Connection
Private strConexao As String
Private strErrorHandler As String

Private mFso As FileSystemObject
Private colArquivos As Collection

Public Enum TiposBase
   Funcionarios = 1
   Malotes = 2
   Descricao = 3
   FunionariosMalotesDescricao = 4
End Enum


Public Property Get Erro() As String
    Erro = strErrorHandler
End Property

Private Sub Class_Initialize()
    Set conexaoComBancoDeDados = New ADODB.Connection
    Set mFso = New FileSystemObject

    strConexao = _
        "Driver={PostgreSQL ANSI};" & _
        "Server=localhost;" & _
        "Port=5432;" & _
        "Database=projects;" & _
        "Uid=postgres;" & _
        "Pwd=Ss_123!;"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next

    If Not conexaoComBancoDeDados Is Nothing Then
        If conexaoComBancoDeDados.State = adStateOpen Then conexaoComBancoDeDados.Close
    End If

    Set conexaoComBancoDeDados = Nothing
    Set mFso = Nothing
    Set colArquivos = Nothing
End Sub

Public Function AtualizarBases(diretorioFinal As String, conteudoArquivo As Integer) As Boolean
   On Error GoTo AtualizarBases_E
   
   AtualizarBases = False
   
   If Not AbrirConexao Then
      GoTo AtualizarBases_E
   End If
   
   If Not (ListarArquivos(diretorioFinal, conteudoArquivo)) Then
      GoTo AtualizarBases_E
   End If
   
   If Not LerArquivos Then
      GoTo AtualizarBases_E
   End If
   
   AtualizarBases = True
    
   Exit Function

AtualizarBases_E:
   strErrorHandler = Err.Description
   FecharConexao
End Function

Private Function AbrirConexao() As Boolean
   On Error GoTo AbrirConexao_E
   
   AbrirConexao = False
   
   If conexaoComBancoDeDados.State <> adStateOpen Then
      conexaoComBancoDeDados.Open strConexao
   End If
   
   AbrirConexao = True
   
   Exit Function
   
AbrirConexao_E:
   strErrorHandler = Err.Description
End Function

Private Sub FecharConexao()
    If conexaoComBancoDeDados.State = adStateOpen Then
        conexaoComBancoDeDados.Close
    End If
End Sub

Private Function ListarArquivos(pasta As String, conteudoArquivo As Integer) As Boolean
   On Error GoTo ListarArquivos_E
   
   ListarArquivos = False
   
   Dim fso As Object
   Dim pastaObj As Object
   Dim Arquivo As Object
   Dim caminhoArq As String
   
   pasta = Left(pasta, 35)
   
   Set colArquivos = New Collection
   
   Set fso = CreateObject("Scripting.FileSystemObject")

   Set pastaObj = fso.GetFolder(pasta)
   
   For Each Arquivo In pastaObj.Files
      If conteudoArquivo = FunionariosMalotesDescricao Then
         caminhoArq = Arquivo.Path
         colArquivos.Add caminhoArq
      
      ElseIf conteudoArquivo = Funcionarios And Arquivo.Name = "Funcionario.txt" Then
         caminhoArq = Arquivo.Path
         colArquivos.Add caminhoArq
         Exit For
      
      ElseIf conteudoArquivo = Descricao And Arquivo.Name = "Descricao.txt" Then
         caminhoArq = Arquivo.Path
         colArquivos.Add caminhoArq
         Exit For
         
      ElseIf conteudoArquivo = Malotes And Arquivo.Name = "Malotes.txt" Then
         caminhoArq = Arquivo.Path
         colArquivos.Add caminhoArq
         Exit For
         
      End If
   Next Arquivo
   
   ListarArquivos = True
   
   GoTo DestruirObjetos
   
ListarArquivos_E:
   strErrorHandler = Err.Description
   
DestruirObjetos:
   Set pastaObj = Nothing
   Set fso = Nothing
End Function

Private Function LerArquivos() As Boolean
   On Error GoTo LerArquivos_E
   
   LerArquivos = False
   
   Dim i As Long
   Dim arq As TextStream
   Dim colLinhas As Collection
   Dim strLinha As String
   
   For i = 1 To colArquivos.Count
   
      Set colLinhas = New Collection
      Set arq = mFso.OpenTextFile(colArquivos(i), ForReading)
   
      Do While Not arq.AtEndOfStream
         strLinha = arq.ReadLine
         colLinhas.Add strLinha
      Loop
   
      arq.Close
      
      If Not (ProcessarArquivo(colLinhas)) Then
         GoTo LerArquivos_E
      End If
   
      Set colLinhas = Nothing
   Next
   
   LerArquivos = True
   
   GoTo DestruirObjetos
   
LerArquivos_E:
   strErrorHandler = Err.Description
   
DestruirObjetos:
   Set colArquivos = Nothing
   Set colLinhas = Nothing
   Set arq = Nothing

End Function

Private Function ProcessarArquivo(ByVal colLinhas As Collection) As Boolean
   On Error GoTo ProcessarArquivo_E
   
   ProcessarArquivo = False
  
   Dim cabecalho As String
   Dim i As Long
   
   If colLinhas.Count = 0 Then
      Exit Function
   End If
   
   cabecalho = colLinhas(1)
   
   conexaoComBancoDeDados.BeginTrans
   
   For i = 2 To colLinhas.Count
      If Not (InserirLinha(cabecalho, colLinhas(i))) Then
         GoTo ProcessarArquivo_E
      End If
   Next
   
   conexaoComBancoDeDados.CommitTrans
   
   ProcessarArquivo = True
   
   Exit Function

ProcessarArquivo_E:
   If conexaoComBancoDeDados.State = adStateOpen Then
      conexaoComBancoDeDados.RollbackTrans
      conexaoComBancoDeDados.Close
      strErrorHandler = Err.Description & " Erro PostgreSQL"
   End If
   
End Function

Private Function InserirLinha(ByVal cabecalho As String, ByVal linha As String) As Boolean
   On Error GoTo InserirLinha_E
   
   InserirLinha = False
   
   Dim colCampos As Collection
   Set colCampos = SplitLinha(linha)

   Select Case cabecalho
   Case "Código;Descrição"
      If Not (InserirDescricao(colCampos)) Then
         GoTo InserirLinha_E
      End If
   
   Case "Código;Matricula;Nome;Data Nascimento"
      If Not (InserirFuncionario(colCampos)) Then
         GoTo InserirLinha_E
      End If
   
   Case "Código;Código Funcionário;Data de Envio;Data Conferencia;Situacao Malote;Código da Descrição"
      If Not (InserirMalote(colCampos)) Then
         GoTo InserirLinha_E
      End If
   
   End Select
    
   InserirLinha = True
    
   GoTo DestruirObjetos
    
InserirLinha_E:
   strErrorHandler = Err.Description

DestruirObjetos:
    Set colCampos = Nothing
End Function

Private Function InserirDescricao(ByVal c As Collection) As Boolean
   On Error GoTo InserirDescricao_E
   
   InserirDescricao = False
   
   Dim strSQL As String
   
   strSQL = "INSERT INTO SUPDESCRICAO (SEQCODDESCRICAO, DESCRICAO) VALUES (nextval('SEQCODDESCRICAO'), " & SqlTexto(c(2)) & ")"
   
   If Not (ExecutarSQL(strSQL)) Then
      GoTo InserirDescricao_E
   End If
        
   InserirDescricao = True
   
   Exit Function
        
InserirDescricao_E:
   strErrorHandler = Err.Description
End Function

Private Function InserirFuncionario(ByVal c As Collection) As Boolean
   On Error GoTo InserirFuncionario_E
   
   InserirFuncionario = False
   
   Dim strSQL As String
   
   strSQL = "INSERT INTO SUPFUNCIONARIO (SEQCODFUNCIONARIO, MATRICULA, NOME, DTNASC) " & _
            "VALUES (nextval('SEQCODFUNCIONARIO'), " & SqlNumero(c(2)) & ", " & SqlTexto(c(3)) & ", " & SqlData(c(4)) & ")"
   
   If Not (ExecutarSQL(strSQL)) Then
      GoTo InserirFuncionario_E
   End If
   
   InserirFuncionario = True
   
   Exit Function
   
InserirFuncionario_E:
   strErrorHandler = Err.Description
End Function

Private Function InserirMalote(ByVal c As Collection) As Boolean
   On Error GoTo InserirMalote_E
   
   InserirMalote = False
   
   Dim strSQL As String
   
   strSQL = "INSERT INTO SUPMALOTE (SEQCODMALOTE, CODFUNCIONARIO,DTENVIO,DTCONFERENCIA,SITUACAOMALOTE,CODDESCRICAO) " & _
            "VALUES (nextval('SEQCODMALOTE'), " & SqlNumero(c(2)) & ", " & SqlData(c(3)) & ", " & SqlData(c(4)) & ", " & SqlTexto(c(5)) & ", " & SqlNumero(c(6)) & ")"
   
   If Not (ExecutarSQL(strSQL)) Then
      GoTo InserirMalote_E
   End If
      
   InserirMalote = True
   
   Exit Function
       
InserirMalote_E:
   strErrorHandler = Err.Description
End Function

Private Function ExecutarSQL(ByVal strSQL As String) As Boolean
   On Error GoTo ExecutarSQL_E
   
   ExecutarSQL = False
   
   conexaoComBancoDeDados.Execute strSQL, , adExecuteNoRecords
   
   ExecutarSQL = True
   
   Exit Function
   
ExecutarSQL_E:
   If conexaoComBancoDeDados.State = adStateOpen Then
      conexaoComBancoDeDados.RollbackTrans
      conexaoComBancoDeDados.Close
      strErrorHandler = Err.Description & " Erro PostgreSQL"
   End If
End Function

Private Function SplitLinha(ByVal linha As String) As Collection
    Dim col As New Collection
    Dim buffer As String
    Dim i As Long
    Dim ch As String

    buffer = ""

    For i = 1 To Len(linha) + 1
        If i <= Len(linha) Then
            ch = Mid(linha, i, 1)
        Else
            ch = ";"
        End If

        If ch = ";" Then
            col.Add buffer
            buffer = ""
        Else
            If ch = "'" Then ch = "''"
            buffer = buffer & ch
        End If
    Next

    Set SplitLinha = col
End Function

Private Function SqlTexto(ByVal v As String) As String
    If Trim(v) = "" Then
        SqlTexto = "NULL"
    Else
        SqlTexto = "'" & Replace(v, "'", "''") & "'"
    End If
End Function

Private Function SqlNumero(ByVal v As String) As String
    If Trim(v) = "" Then
        SqlNumero = "NULL"
    Else
        SqlNumero = v
    End If
End Function

Private Function SqlData(ByVal v As Variant) As String
    Dim dt As Date

    If Trim(CStr(v)) = "" Then
        SqlData = "NULL"
        Exit Function
    End If

    dt = CDate(v)

    ' ISO é o formato mais seguro no PostgreSQL
    SqlData = "'" & Format(dt, "yyyy-mm-dd") & "'"
End Function

