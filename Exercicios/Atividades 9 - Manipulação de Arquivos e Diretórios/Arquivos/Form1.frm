VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Arquivos"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBKP 
      Caption         =   "Bacukup"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "Listar"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton cmdLer 
      Caption         =   "Ler"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "Teste"
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBKP_Click()
  Dim caminhoOrigem As String
  Dim caminhoDestino As String
     
  caminhoOrigem = InputBox("Digite o caminho para realizar o backup.")
  caminhoDestino = InputBox("Digite o caminho de onde você quer salvar o backup.")
  
  Dim OperacaoBKP As clsOperacoesArq
  Set OperacaoBKP = New clsOperacoesArq
  
  If OperacaoBKP.BKP(caminhoOrigem, caminhoDestino) Then
      MsgBox "Backup concluído! Verifique sua pasta " & caminhoDestino & "!"
  Else
      MsgBox "Backup incompleto! Houve erro ao tentar realizar o procedimento."
  End If
End Sub

Private Sub cmdExcluir_Click()
   Dim caminhoArq As String
   caminhoArq = InputBox("Digite o caminho e o nome do arquivo a ser excluído.")
   
   Dim OperacaoExcluir As clsOperacoesArq
   Set OperacaoExcluir = New clsOperacoesArq
   
   If OperacaoExcluir.Excluir(caminhoArq) Then
      MsgBox "Arquivo localizado em " & caminhoArq & " excluído!"
   Else
      MsgBox "Erro ao excluir o arquivo: " & caminhoArq & "!"
   End If
End Sub

Private Sub cmdLer_Click()
   Dim caminhocaminhoPastaComArq As String
   caminhocaminhoPastaComArq = InputBox("Digite o caminho de onde você quer ler o arquivo e o nome do mesmo.")
   
   Dim OperacaoLer As clsOperacoesArq
   Set OperacaoLer = New clsOperacoesArq
   
   If OperacaoLer.Ler(caminhocaminhoPastaComArq) Then
      MsgBox "Deu bom"
   Else
      MsgBox "Deu ruim"
   End If
End Sub

Private Sub cmdListar_Click()
   Dim caminhoPasta As String
   caminhoPasta = InputBox("Digite o caminho de onde você quer listar os arquivos.")
      
   Dim OperacaoListar As clsOperacoesArq
   Set OperacaoListar = New clsOperacoesArq
   
   If (OperacaoListar.ListarPastasSubPastasArquivos(caminhoPasta)) Then
      MsgBox "Arquivos listados, verifique sua tela Immediate!"
   Else
      MsgBox "Houve erro ao listar os arquivos no diretório informado."
   End If
End Sub

Private Sub cmdSalvar_Click()
   Dim caminhoPastaComArquivo As String
   caminhoPastaComArquivo = InputBox("Digite o caminho onde você quer salvar o arquivo.", "Pasta")
    
   Dim nomeArq As String
   nomeArq = InputBox("Digite o nome do arquivo a ser salvo com a respectiva extensão (.txt, .json, .xml).", "Nome")
   
   Dim OperacaoSalvar As clsOperacoesArq
   Set OperacaoSalvar = New clsOperacoesArq
  
   If OperacaoSalvar.Salvar(caminhoPastaComArquivo, nomeArq, Text1.Text) Then
      MsgBox "Deu bom"
   Else
      MsgBox "Deu ruim"
   End If
End Sub
