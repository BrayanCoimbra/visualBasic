VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "V E L H A "
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTabuleiro 
      BackColor       =   &H80000004&
      Caption         =   "TABULEIRO"
      Height          =   2535
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2655
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1560
         Width           =   195
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1560
         Width           =   195
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1560
         Width           =   195
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1080
         Width           =   195
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   5
         Top             =   600
         Width           =   195
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   720
         MaxLength       =   1
         TabIndex        =   4
         Top             =   600
         Width           =   195
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   3
         Top             =   600
         Width           =   195
      End
      Begin VB.Line Line4 
         X1              =   720
         X2              =   1920
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line3 
         X1              =   720
         X2              =   1920
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         X1              =   1560
         X2              =   1560
         Y1              =   600
         Y2              =   1920
      End
      Begin VB.Line Line1 
         X1              =   1080
         X2              =   1080
         Y1              =   600
         Y2              =   1920
      End
   End
   Begin VB.CommandButton cmdJogador2 
      Caption         =   "Jogador 2 (O)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdJogador1 
      Caption         =   "Jogador 1 (X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim colTextBoxes As Collection

Private Sub Form_Load()
    Set colTextBoxes = New Collection
    ' Adicionar caixas de texto à coleção
    colTextBoxes.Add Text1, "Text1"
    colTextBoxes.Add Text2, "Text2"
    colTextBoxes.Add Text3, "Text3"
    colTextBoxes.Add Text4, "Text4"
    colTextBoxes.Add Text5, "Text5"
    colTextBoxes.Add Text6, "Text6"
    colTextBoxes.Add Text7, "Text7"
    colTextBoxes.Add Text8, "Text8"
    colTextBoxes.Add Text9, "Text9"
End Sub

Private Sub cmdJogador1_Click()
    Dim guardaLinhaJogador1 As Integer
    Dim guardaColunaJogador1 As Integer
    
    guardaLinhaJogador1 = InputBox("Em qual linha você quer marcar?", "Jogador 1", "")
    guardaColunaJogador1 = InputBox("Em qual coluna você quer marcar?", "Jogador 1", "")
    
    If guardaLinhaJogador1 > 0 And guardaLinhaJogador1 < 4 And guardaColunaJogador1 > 0 And guardaColunaJogador1 < 4 Then
        colTextBoxes.Item("Text" & ((guardaLinhaJogador1 - 1) * 3 + guardaColunaJogador1)).Text = "X"
    Else
        MsgBox "Você pode digitar apenas nas linhas e colunas entre os valores 1 e 3.", vbExclamation, "ATENÇÃO"
    End If
    
    'Verificar se há condições para vencer na colunas
    If colTextBoxes("Text1").Text = "X" And colTextBoxes("Text2").Text = "X" And colTextBoxes("Text3").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text4").Text = "X" And colTextBoxes("Text5").Text = "X" And colTextBoxes("Text6").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text7").Text = "X" And colTextBoxes("Text8").Text = "X" And colTextBoxes("Text9").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    
    'Verificar se há condições para vencer na colunas
    ElseIf colTextBoxes("Text1").Text = "X" And colTextBoxes("Text4").Text = "X" And colTextBoxes("Text7").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text2").Text = "X" And colTextBoxes("Text5").Text = "X" And colTextBoxes("Text8").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text3").Text = "X" And colTextBoxes("Text6").Text = "X" And colTextBoxes("Text9").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
        
    'Verifica se há condições para vencer nas diagonais
    ElseIf colTextBoxes("Text1").Text = "X" And colTextBoxes("Text5").Text = "X" And colTextBoxes("Text9").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text3").Text = "X" And colTextBoxes("Text5").Text = "X" And colTextBoxes("Text7").Text = "X" Then
        msg = MsgBox("Jogador 1 Ganhou!", vbExclamation, "Parabéns")
    End If
    
End Sub

Private Sub cmdJogador2_Click()
    Dim guardaLinhaJogador2 As Integer
    Dim guardaColunaJogador2 As Integer
    
    guardaLinhaJogador2 = InputBox("Em qual linha você quer marcar?", "Jogador 2", "")
    guardaColunaJogador2 = InputBox("Em qual coluna você quer marcar?", "Jogador 2", "")
    
    If guardaLinhaJogador2 > 0 And guardaLinhaJogador2 < 4 And guardaColunaJogador2 > 0 And guardaColunaJogador2 < 4 Then
        colTextBoxes.Item("Text" & ((guardaLinhaJogador2 - 1) * 3 + guardaColunaJogador2)).Text = "O"
    Else
        MsgBox "Você pode digitar apenas nas linhas e colunas entre os valores 1 e 3.", vbExclamation, "ATENÇÃO"
    End If
    
    'Verificar se há condições para vencer na colunas
    If colTextBoxes("Text1").Text = "O" And colTextBoxes("Text2").Text = "O" And colTextBoxes("Text3").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text4").Text = "O" And colTextBoxes("Text5").Text = "O" And colTextBoxes("Text6").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text7").Text = "O" And colTextBoxes("Text8").Text = "O" And colTextBoxes("Text9").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    
    'Verificar se há condições para vencer na colunas
    ElseIf colTextBoxes("Text1").Text = "O" And colTextBoxes("Text4").Text = "O" And colTextBoxes("Text7").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text2").Text = "O" And colTextBoxes("Text5").Text = "O" And colTextBoxes("Text8").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text3").Text = "O" And colTextBoxes("Text6").Text = "O" And colTextBoxes("Text9").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
        
    'Verifica se há condições para vencer nas diagonais
    ElseIf colTextBoxes("Text1").Text = "O" And colTextBoxes("Text5").Text = "O" And colTextBoxes("Text9").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    ElseIf colTextBoxes("Text3").Text = "O" And colTextBoxes("Text5").Text = "O" And colTextBoxes("Text7").Text = "O" Then
        msg = MsgBox("Jogador 2 Ganhou!", vbExclamation, "Parabéns")
    End If
End Sub

