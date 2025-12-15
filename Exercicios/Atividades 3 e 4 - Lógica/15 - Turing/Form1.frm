VERSION 5.00
Object = "{263D3036-6BF5-11D5-A656-0080C8BAEF42}#1.4#0"; "LydiansEdit.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ler String"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin LydiansOcx.txt txt2 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   873
      TitOculto       =   -1  'True
      TitNome         =   "Resultado"
   End
   Begin LydiansOcx.txt txt1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Turing As New clsTuring

Private Sub Command1_Click()
   Turing.Iniciar txt1.Text
   txt2.Text = Turing.gResultado
End Sub

Private Sub Form_Load()
   Turing.lEstado = "O"
End Sub
