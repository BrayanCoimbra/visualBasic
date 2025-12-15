VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000F9100&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " C R A P S"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   FillColor       =   &H000F9100&
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000F9100&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   7515
   ScaleWidth      =   10155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDados 
      BackColor       =   &H000F9100&
      Caption         =   "D A D O S"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   2880
      TabIndex        =   1
      Top             =   3360
      Width           =   6975
      Begin VB.PictureBox picBox2 
         BackColor       =   &H000F9100&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   3720
         ScaleHeight     =   2835
         ScaleWidth      =   2835
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.PictureBox picBox1 
         BackColor       =   &H000F9100&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   360
         ScaleHeight     =   2835
         ScaleWidth      =   2955
         TabIndex        =   2
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdLancarDados 
      Appearance      =   0  'Flat
      BackColor       =   &H000F9100&
      Caption         =   "Lançar Dados! "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaskColor       =   &H000F9100&
      TabIndex        =   0
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000F9100&
      Caption         =   "Seu ponto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000F9100&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9480
      TabIndex        =   4
      Top             =   3120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Craps As New clsCraps

Private Sub cmdLancarDados_Click()
  
   Craps.LancarDados
   
   picBox1.Picture = LoadPicture(Craps.gCaminhoDado1)
   picBox2.Picture = LoadPicture(Craps.gCaminhoDado2)
   Label1.Caption = Craps.gPonto
   
   ' Verificar as condições de vitória na primeira jogada
   If Craps.gNumeroDeExecucoes = 2 Then
      If Craps.gSoma = 11 Or Craps.gSoma = 7 Then
         MsgBox Craps.gMensagem, vbInformation, "Parabéns"
         Craps.lNumeroDeExecucoes = 1
         Craps.lDado1 = 0
         Craps.lDado2 = 0
         Craps.lPonto = 0
         Craps.lSoma = 0
         Label1.Caption = ""
      End If
   End If
         
   ' Verificar as condições de derrota
   If Craps.gSoma = 2 Or Craps.gSoma = 3 Or Craps.gSoma = 12 Then
      MsgBox "Craps! Você perdeu.", vbExclamation, "Ops!"
      Craps.lNumeroDeExecucoes = 1
      Craps.lDado1 = 0
      Craps.lDado2 = 0
      Craps.lPonto = 0
      Craps.lSoma = 0
      Label1.Caption = ""
   End If
      
   ' Verificar as condições de vitória com Ponto a partir da segunda jogada, visto que a primeira poderá guardar o ponto
   If Craps.gNumeroDeExecucoes > 2 And Craps.gPonto = Craps.gSoma Then
      If Craps.gPonto = 4 Or Craps.gPonto = 5 Or Craps.gPonto = 6 Or Craps.gPonto = 8 Or Craps.gPonto = 9 Or Craps.gPonto = 10 Then
         MsgBox Craps.gMensagem, vbInformation, "Parabéns!"
         Craps.lNumeroDeExecucoes = 1
         Craps.lDado1 = 0
         Craps.lDado2 = 0
         Craps.lPonto = 0
         Craps.lSoma = 0
         Label1.Caption = ""
      End If
   End If
   
   'Verifica possibilidade de derroa se os dados somarem 7 após a segunda jogada
   If Craps.gNumeroDeExecucoes > 1 Then
      If Craps.gSoma = 7 Then
         MsgBox Craps.gMensagem, vbExclamation, "Ops!"
         Craps.lNumeroDeExecucoes = 1
         Craps.lDado1 = 0
         Craps.lDado2 = 0
         Craps.lPonto = 0
         Craps.lSoma = 0
         Label1.Caption = ""
      End If
   End If
End Sub

Private Sub Form_Load()
   Craps.lNumeroDeExecucoes = 1
End Sub


'Após os dados pararem, a soma dos pontos das duas faces para cima são somadas.

'Se a soma for 7 ou 11 no primeiro lançamento o jogador ganha.
'se a soma for 2, 3 ou 12 (chamado de Crap) o jogador perde.
'Se a soma for 4, 5, 6, 8, 9 ou 10 no primeiro lançamento é chamado de "ponto" do jogador.

'Para ganhar o jogador deve lançar novamente os dados até atingir o "ponto" (a soma dos dados do primeiro lançamento).
'Caso a soma desse novo lançamento for 7 o jogador perde.
'Isto ocorre até o valor do ponto ser repetido ou o 7 aparecer.
