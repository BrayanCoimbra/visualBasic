VERSION 5.00
Object = "{263D3036-6BF5-11D5-A656-0080C8BAEF42}#1.5#0"; "LydiansEdit.ocx"
Begin VB.Form Form1 
   Caption         =   "Days Lived"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   1620
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Data Atual"
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3375
      Begin LydiansOcx.vlr vlrFimDia 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         ForeColor       =   -2147483631
         Enabled         =   0   'False
         TitNome         =   "Dia:"
         TitLargura      =   346
         Decimais        =   0
         Inteiros        =   2
      End
      Begin LydiansOcx.vlr vlrFimMes 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         ForeColor       =   -2147483631
         Enabled         =   0   'False
         TitNome         =   "Mês:"
         TitLargura      =   406
         Decimais        =   0
         Inteiros        =   2
      End
      Begin LydiansOcx.vlr vlrFimAno 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Milhar          =   0   'False
         ForeColor       =   -2147483631
         Enabled         =   0   'False
         TitNome         =   "Ano:"
         TitLargura      =   391
         Decimais        =   0
         Inteiros        =   4
      End
   End
   Begin VB.Frame fraDtIni 
      Caption         =   "Data Nasc."
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   3375
      Begin LydiansOcx.vlr vlrIniDia 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   556
         ForeColor       =   0
         TitNome         =   "Dia:"
         TitLargura      =   346
         Decimais        =   0
         Inteiros        =   2
         Value           =   19
      End
      Begin LydiansOcx.vlr vlrIniMes 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   556
         ForeColor       =   0
         TitNome         =   "Mês:"
         TitLargura      =   406
         Decimais        =   0
         Inteiros        =   2
         Value           =   2
      End
      Begin LydiansOcx.vlr vlrIniAno 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         Milhar          =   0   'False
         ForeColor       =   0
         TitNome         =   "Ano:"
         TitLargura      =   391
         Decimais        =   0
         Inteiros        =   4
         Value           =   2001
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CalculadoraDeDias As CalculadoraDeDias

Private Sub Command1_Click()
   With CalculadoraDeDias
      If (.Inicializar(Me.vlrIniDia, Me.vlrIniMes, Me.vlrIniAno)) Then
         MsgBox "Dias vividos desde a data de nascimento: " & .DiasVividos, vbInformation, "Days lived"
      Else
         MsgBox "Erro inesperado. Contate o programador.", vbExclamation, "Erro"
      End If
   End With
End Sub

Private Sub Form_Load()
   Set CalculadoraDeDias = New CalculadoraDeDias
   
   dataAtual = Date
   diaAtual = Day(dataAtual)
   mesAtual = Month(dataAtual)
   anoAtual = Year(dataAtual)

   Me.vlrFimDia.Value = diaAtual
   Me.vlrFimMes.Value = mesAtual
   Me.vlrFimAno.Value = anoAtual
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set CalculadoraDeDias = Nothing
End Sub
