VERSION 5.00
Begin VB.Form frmSEDEXFuncsMalotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Funcionários"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6375
   Icon            =   "frmSEDEXFuncsMalotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNmFunc 
      Height          =   285
      Left            =   240
      TabIndex        =   14
      Top             =   960
      Width           =   4455
   End
   Begin VB.TextBox dtNascFunc 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox vlrMatriculaFunc 
      Height          =   285
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdAtualizarBase 
      Caption         =   "&Atualizar Base"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Frame fraMalotes 
      Caption         =   "Malotes"
      Height          =   2400
      Left            =   60
      TabIndex        =   7
      Top             =   1800
      Width           =   4755
      Begin VB.TextBox txtMalotes 
         Height          =   2055
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame fraIdentificacao 
      Caption         =   "Identificação"
      Height          =   615
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   4755
      Begin VB.TextBox vlrCod 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2220
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdConsultar 
      Caption         =   "&Consultar"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   540
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1380
      Width           =   1335
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame fraParametros 
      Caption         =   "Funcionário"
      Enabled         =   0   'False
      Height          =   975
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   4755
   End
   Begin VB.PictureBox imgIcones 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   7440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   3480
      Width           =   1200
   End
End
Attribute VB_Name = "frmSEDEXFuncsMalotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DataAccessLayer As clsDataAccessLayer
Private colFuncionarios As Collection
Private colMalotes As Collection

Private Enum Funcionario
   Nome = 1
   Matricula = 2
   DataDeNascimento = 3
End Enum

Private Enum Malotes
   CodMalote = 0
   DataEnvioDoMalote = 1
   situacaomalote = 2
   DescricaoSituacao = 3
End Enum

Private Sub cmdAtualizarBase_Click()
   On Error GoTo cmdAtualizarBase_Click_E
   
   Dim frmAtualizarBase As frmSEDEXAtualizarBase
   Set frmAtualizarBase = New frmSEDEXAtualizarBase
   
   frmAtualizarBase.Show vbModal
   
   GoTo DestruirObjetos
   
cmdAtualizarBase_Click_E:
   MsgBox Err.Description
   
DestruirObjetos:
   Set frmAtualizarBase = Nothing
End Sub

Private Sub cmdConsultar_Click()
   On Error GoTo cmdConsultar_Click_E
      
   Dim ItemX As ListItem
   Dim i As Integer
   
  'Validar valor vazio, se for vazio, consultará todos os registros
   If Not IsNumeric(Me.vlrCod) Then
      MsgBox "Informe corretamente o código!"
      Exit Sub
   End If
   
   If Not Me.vlrCod > 0 Then
      MsgBox "Informe um código válido!"
      Exit Sub
   End If
   
   With DataAccessLayer
      If Not .ConsultarFuncionarioMalote(Str(Me.vlrCod)) Then
         MsgBox .getErrorHandler
      End If
      
      Set colFuncionarios = .getColFuncionarios
      Set colMalotes = .getColMalotes
      
      For i = 1 To colMalotes.Count Step 4
         If colMalotes(i) <> "" Then
            Me.txtMalotes = Me.txtMalotes & "Cód. Malote: " & colMalotes(i) & " - " 'CodMalote
            Me.txtMalotes = Me.txtMalotes & "Data de Envio do Malote: " & colMalotes(i + 1) & " - " 'DataEnvioDoMalote
            Me.txtMalotes = Me.txtMalotes & "Situação do Malote: " & colMalotes(i + 2) & " - " 'SituacaoMalote
            Me.txtMalotes = Me.txtMalotes & "Descrição: " & colMalotes(i + 3) & vbNewLine 'DescricaoSituacao
         End If
      Next i
      
      Me.cmdConsultar.Enabled = False
      Me.cmdSalvar.Enabled = True
      Me.fraParametros.Enabled = True
      Me.fraIdentificacao.Enabled = False
      Me.cmdExcluir.Enabled = True
      Me.cmdSalvar.Caption = "&Alterar"
      
      Me.txtNmFunc = colFuncionarios(Nome) 'NomeFuncionario
      Me.vlrMatriculaFunc = colFuncionarios(Matricula) 'Matricula
      Me.dtNascFunc = colFuncionarios(DataDeNascimento) 'Data de Nascimento
            
      If Not .LimparCollections Then
         MsgBox .getErrorHandler
      End If
   End With
     
   GoTo DestruirObjetos
   
cmdConsultar_Click_E:
   MsgBox Err.Description
   
DestruirObjetos:
   Set colFuncionarios = Nothing
   Set colMalotes = Nothing
End Sub

Private Sub cmdExcluir_Click()
   On Error GoTo cmdExcluir_Click_E
   
   Dim intConfirm As Integer
   
   intConfirm = MsgBox("Você perderá as informações de todos os malotes vinculados a este funcionário se continuar. Tem certeza que quer prosseguir?", vbYesNo + vbQuestion, "Confirmação")
   
   If intConfirm = vbYes Then
   
      With DataAccessLayer
        If Not .ExcluirFuncionarioMalote(Str(Me.vlrCod)) Then
          MsgBox .getErrorHandler
        End If
      End With
   
      MsgBox "Excluído com sucesso!"
   
   Else
      MsgBox "Exclusão cancelada!"
   End If

   cmdLimpar_Click
  
   GoTo DestruirObjetos
   
cmdExcluir_Click_E:
   MsgBox Err.Description
   
DestruirObjetos:
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdLimpar_Click()
   On Error GoTo cmdLimpar_Click_E

   Me.vlrCod = ""
   
   Me.cmdSalvar.Caption = "&Salvar"
   
   Me.fraParametros.Enabled = False
   Me.fraIdentificacao.Enabled = True
   Me.txtNmFunc = ""
   Me.vlrMatriculaFunc = ""
   Me.dtNascFunc = ""
   Me.txtMalotes = ""
    
   Me.cmdConsultar.Enabled = True
   Me.cmdSalvar.Enabled = False
   Me.cmdExcluir.Enabled = False
   
    If Not DataAccessLayer.LimparCollections Then
      MsgBox DataAccessLayer.getErrorHandler
   End If

   GoTo DestruirObjetos
   
cmdLimpar_Click_E:
   MsgBox Err.Description
   
DestruirObjetos:
End Sub

Private Sub cmdSalvar_Click()
   On Error GoTo cmdSalvar_Click_E
   
    Dim strNome As String
    Dim strUpdateNome As String
    Dim intConfirm As Integer
    Dim i As Integer
   
    intConfirm = MsgBox("Confirmar a atualização dos dados?", vbYesNo + vbQuestion, "Confirmação")
   
    If Not intConfirm = vbYes Then
       MsgBox "Atualização de dados cancelada!"
    End If
   
    strNome = ""
    strNome = Me.txtNmFunc
   
    If Me.txtNmFunc <> "" And Me.vlrMatriculaFunc > 0 And Me.dtNascFunc <> "" Then
       
       For i = 1 To Len(strNome)
          If Asc(Mid(strNome, i, 1)) >= Asc("a") And Asc(Mid(strNome, i, 1)) <= Asc("z") Or _
             Asc(Mid(strNome, i, 1)) >= Asc("A") And Asc(Mid(strNome, i, 1)) <= Asc("Z") Or _
             Asc(Mid(strNome, i, 1)) = Asc("'") Or Asc(Mid(strNome, i, 1)) = Asc(" ") Or Asc(Mid(strNome, i, 1)) = Asc(".") Then
                
             strUpdateNome = strUpdateNome & Mid(strNome, i, 1)
               
             If Asc(Mid(strNome, i, 1)) = Asc("'") Then
                strUpdateNome = strUpdateNome & Mid(strNome, i, 1)
             End If
          Else
             MsgBox "O campo - Nome - está impróprio!"
             GoTo DestruirObjetos
          End If
       Next i
      
    Else
       MsgBox "Preencha corretamente os campos!"
      
       GoTo DestruirObjetos
    End If

   With DataAccessLayer
     If Not .AtualizarFuncionarioMalote(Str(Me.vlrCod), strUpdateNome, Str(vlrMatriculaFunc), Me.dtNascFunc) Then
       MsgBox .getErrorHandler
     End If
   End With
   
    MsgBox "Informações atualizadas com sucesso!"

    cmdLimpar_Click
  
    GoTo DestruirObjetos
   
cmdSalvar_Click_E:
DestruirObjetos:
End Sub

Private Sub Form_Load()
   Set DataAccessLayer = New clsDataAccessLayer
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set DataAccessLayer = Nothing
End Sub
