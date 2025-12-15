VERSION 5.00
Object = "{263D3036-6BF5-11D5-A656-0080C8BAEF42}#1.5#0"; "LydiansEdit.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar"
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin LydiansOcx.vlr vlrP 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Gasto - Paulo"
      TitLargura      =   1024
      Decimais        =   0
      Value           =   90
   End
   Begin LydiansOcx.vlr vlrJP 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Lucro:"
      TitLargura      =   1024
      Decimais        =   0
      Value           =   630
   End
   Begin LydiansOcx.vlr vlrJ 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      ForeColor       =   0
      TitNome         =   "Gasto - João"
      TitLargura      =   1024
      Decimais        =   0
      Value           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'05- João e Paulo receberam juntos, R$ 630,00 para consertar uns computadores.
'Se João gastar R$ 60,00 do que recebeu, e Paulo gastar R$ 90,00, ambos ficarão com quantias
'iguais. Qual foi a quantia recebida por Paulo? Faça um programa para descobrir quanto cada um recebeu.

Private GastoJoao As Double
Private GastoPaulo As Double
Private QuantiaJoao As Double
Private QuantiaPaulo As Double
Private GastoTotal As Double

Private Sub Command1_Click()
   DescobrirQuantia Me.vlrJ, Me.vlrP, 630
   MsgBox "João recebeu R$ " & QuantiaJoao & vbCrLf & "Paulo recebeu R$ " & QuantiaPaulo, vbInformation, "Quantias Recebidas"
End Sub

Private Sub DescobrirQuantia(pGastoJoao As Double, pGastoPaulo As Double, pGastoTotal As Double)
   ' Gastos de João e Paulo
   GastoJoao = pGastoJoao
   GastoPaulo = pGastoPaulo
   GastoTotal = pGastoTotal
   
   ' Resolvendo o sistema de equações
   QuantiaJoao = (GastoTotal + pGastoPaulo - pGastoJoao) / 2
   QuantiaPaulo = GastoTotal - QuantiaJoao
End Sub


