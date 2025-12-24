VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataAccessLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private conexaoComBancoDeDados As ADODB.Connection
Private conecString As String
Private errorHandler As String
Private colFuncionarios As Collection
Private colMalotes As Collection

Private Enum Funcionario
   Nome = 0
   Matricula = 1
   DtNasc = 2
End Enum

Private Enum Malotes
   CodMalote = 3
   DataEnvioDoMalote = 4
   situacaomalote = 5
   DescricaoSituacao = 6
End Enum

Public Property Get getColFuncionarios() As Collection
   Set getColFuncionarios = colFuncionarios
End Property

Public Property Get getColMalotes() As Collection
   Set getColMalotes = colMalotes
End Property

Public Property Get getErrorHandler() As String
    getErrorHandler = errorHandler
End Property

Public Function LimparCollections() As Boolean
   On Error GoTo LimparCollections_E
   
   LimparCollections = False
   
   Set colFuncionarios = Nothing
   
   Set colMalotes = Nothing
   
   LimparCollections = True
   
   GoTo DestruirObjetos
   
LimparCollections_E:
    errorHandler = Err.Description
   
DestruirObjetos:
End Function

Public Function ConsultarFuncionarioMalote(strPkFuncionarioMalote As String) As Boolean
   On Error GoTo ConsultarFuncionarioMalote_E
   
   ConsultarFuncionarioMalote = False
   
   Dim strQuerie As String
   Dim recordSet As New ADODB.recordSet
   
   Set colFuncionarios = New Collection
   Set colMalotes = New Collection
   
   With conexaoComBancoDeDados
      .Open conecString
      .BeginTrans
   
      With recordSet
      
         strQuerie = " SELECT NOME AS NomeFuncionario, MATRICULA AS Matricula, DTNASC AS DtNasc,  " & _
                  " SUPMALOTE.SEQCODMALOTE AS CodMalote, SUPMALOTE.DTENVIO AS DataEnvioDoMalote, " & _
                  " SUPMALOTE.SITUACAOMALOTE AS SituacaoMalote, SUPDESCRICAO.DESCRICAO AS DescricaoSituacao " & _
                  " FROM SUPMALOTE " & _
                  " JOIN SUPFUNCIONARIO ON SUPMALOTE.CODFUNCIONARIO = SUPFUNCIONARIO.SEQCODFUNCIONARIO " & _
                  " JOIN SUPDESCRICAO ON SUPMALOTE.CODDESCRICAO = SUPDESCRICAO.SEQCODDESCRICAO "
        
         If strPkFuncionarioMalote <> "" Then
            strQuerie = strQuerie & " WHERE SUPFUNCIONARIO.SEQCODFUNCIONARIO = " & strPkFuncionarioMalote
         End If
        
         .Open strQuerie, conexaoComBancoDeDados, adOpenForwardOnly, adLockReadOnly
        
         While Not .EOF
            colFuncionarios.Add .Fields(Nome).Value
            colFuncionarios.Add .Fields(Matricula).Value
            colFuncionarios.Add .Fields(DtNasc).Value
            colMalotes.Add .Fields(CodMalote).Value
            colMalotes.Add .Fields(DataEnvioDoMalote).Value
            colMalotes.Add .Fields(situacaomalote).Value
            colMalotes.Add .Fields(DescricaoSituacao).Value
            recordSet.MoveNext
         Wend
      
      .Close
      
      End With

      .CommitTrans
      .Close
   End With
      
   ConsultarFuncionarioMalote = True

   GoTo DestruirObjetos

ConsultarFuncionarioMalote_E:
   If conexaoComBancoDeDados.State = adStateOpen Then
      conexaoComBancoDeDados.RollbackTrans
     conexaoComBancoDeDados.Close
      errorHandler = Err.Description & " Erro PostgreSQL"
   End If

DestruirObjetos:
End Function

Public Function ExcluirFuncionarioMalote(strPkFuncionarioMalote As String) As Boolean
   On Error GoTo ExcluirFuncionarioMalote_E
   
   ExcluirFuncionarioMalote = False
   
   Dim strQuerieDelete As String
 
   With conexaoComBancoDeDados
      .Open conecString
      .BeginTrans
   
      strQuerieDelete = "DELETE FROM SUPMALOTE WHERE CODFUNCIONARIO = " & strPkFuncionarioMalote

      .Execute strQuerie, , adExecuteNoRecords
   
      strQurieDelete = "DELETE FROM SUPFUNCIONARIO WHERE SEQCODFUNCIONARIO = " & strPkFuncionarioMalote
   
      .Execute strQuerieDelete, , adExecuteNoRecords
   
      .CommitTrans
      .Close
   End With
   
   ExcluirFuncionarioMalote = True
   
   GoTo DestruirObjetos

ExcluirFuncionarioMalote_E:
   If conexaoComBancoDeDados.State = adStateOpen Then
      conexaoComBancoDeDados.RollbackTrans
      conexaoComBancoDeDados.Close
      errorHandler = Err.Description & " Erro PostgreSQL"
   End If

DestruirObjetos:
End Function

Public Function AtualizarFuncionarioMalote(strPkFuncionarioMalote As String, strUpdateNome As String, vlrMatriculaFunc As String, dtNascFunc As String) As Boolean
   On Error GoTo ExcluirFuncionarioMalote_E
   
   AtualizarFuncionarioMalote = False
   
   Dim strQuerieUpdate As String
 
   With conexaoComBancoDeDados
      .Open conecString
      .BeginTrans
     
      strQuerieUpdate = " UPDATE SUPFUNCIONARIO SET " & _
                     " NOME = " & "'" & strUpdateNome & "'" & _
                     ", MATRICULA = " & vlrMatriculaFunc & _
                     ", DTNASC = " & "'" & dtNascFunc & "'" & _
                     " WHERE SEQCODFUNCIONARIO = " & strPkFuncionarioMalote
   
      .Execute strQuerieUpdate, , adExecuteNoRecords
   
      .CommitTrans
      .Close
   End With
   
   AtualizarFuncionarioMalote = True
   
   GoTo DestruirObjetos

ExcluirFuncionarioMalote_E:
   If conexaoComBancoDeDados.State = adStateOpen Then
      conexaoComBancoDeDados.RollbackTrans
      conexaoComBancoDeDados.Close
      errorHandler = Err.Description & " Erro PostgreSQL"
   End If

DestruirObjetos:
End Function

Private Sub Class_Initialize()
   Set conexaoComBancoDeDados = New ADODB.Connection
   
   conecString = _
     "Driver={PostgreSQL ANSI};" & _
     "Server=localhost;" & _
     "Port=5432;" & _
     "Database=projects;" & _
     "Uid=postgres;" & _
     "Pwd=Ss_123!;"
End Sub
