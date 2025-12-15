VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOpcoes 
      Caption         =   "Opções"
      Height          =   1935
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2775
      Begin VB.CommandButton cmdAddValorNoBuffer 
         Caption         =   "Adicionar Valor no Buffer"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdRemoverValorDoBuffer 
         Caption         =   "Liberar slots no Buffer"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmdSubstituirValor 
         Caption         =   "Substituir slots no Buffer"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   2295
      End
   End
   Begin VB.TextBox txtAdicionarValorNoBuffer 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame fraBuffer 
      Caption         =   "B U F F E R"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   5415
      Begin VB.TextBox txtResultadoBuffer 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdDefinirTamanhoBuffer 
      Caption         =   "Definir tamanho do Buffer"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtTamanhoBuffer 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Buffer As New clsBuffer

Private Sub cmdDefinirTamanhoBuffer_Click()
   
   If IsNumeric(txtTamanhoBuffer) Then
      If (Buffer.DefinirTamanhoBuffer(CInt(txtTamanhoBuffer))) Then
      msg = MsgBox("Tamanho do buffer definido com " & txtTamanhoBuffer.Text & " espaços!", vbInformation, "Atenção")
      Buffer.AtualizarExibicaoBuffer txtResultadoBuffer
      End If
   Else
      msg = MsgBox("Digite um valor numérico para adicionar no buffer", vbExclamation, "Erro de entrada")
   End If
   
End Sub

Private Sub cmdAddValorNoBuffer_Click()
   
   If IsNumeric(txtAdicionarValorNoBuffer) Then
      If (Buffer.AddValorNoBuffer(CInt(txtAdicionarValorNoBuffer.Text))) Then
         Buffer.AtualizarExibicaoBuffer txtResultadoBuffer
      End If
   Else
    msg = MsgBox("Digite um valor numérico para adicionar no buffer", vbExclamation, "Erro de entrada")
   End If
   
   txtAdicionarValorNoBuffer = ""
     
End Sub

Private Sub cmdRemoverValorDoBuffer_Click()
   
   If (Buffer.RmvValorNoBuffer) Then
      Buffer.AtualizarExibicaoBuffer txtResultadoBuffer
   Else
      msg = MsgBox("Buffer vazio!", vbExclamation, "Atenção")
   End If
  
End Sub

Private Sub cmdSubstituirValor_Click()
   
   If (Buffer.SubstituiValorNoBuffer(txtAdicionarValorNoBuffer.Text)) Then
      msg = MsgBox("Valor substituído com sucesso!", vbInformation, "Sucesso")
   End If
   
   Buffer.AtualizarExibicaoBuffer txtResultadoBuffer
   
End Sub
